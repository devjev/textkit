//! Render Jupyter Notebooks
//!

use crate::render::{
    char_text_tokens, end_tag_event, heading_prequel_tokens, heading_sequel_tokens,
    image_paragraph_tokens, markdown::markdown_to_tokens, monospace_paragraph_tokens,
    paragraph_prequel_tokens, paragraph_sequel_tokens, run_end_token, run_start_token,
    start_tag_event,
};
use crate::{FileContents, ImageFileContents};
use crate::{Token, TokenType, NS_WP_ML};
use pulldown_cmark::{Options, Parser};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};
use std::io::Cursor;

#[derive(Serialize, Deserialize, Debug)]
#[serde(rename_all = "snake_case")]
pub enum JupyterCellType {
    Markdown,
    Code,
    Raw,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct JupyterCellOutput {
    // FIXME this is a crutch, but at least a quick one.
    pub data: HashMap<String, serde_json::Value>,
    pub execution_count: Option<usize>,
    pub output_type: String,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct JupyterCell {
    pub cell_type: JupyterCellType,
    pub source: Vec<String>,
    pub outputs: Option<Vec<JupyterCellOutput>>,
    pub execution_count: Option<usize>,
}

#[derive(Serialize, Deserialize, Debug)]
pub struct JupyterNotebook {
    pub cells: Vec<JupyterCell>,
    pub nbformat: usize,
    pub nbformat_minor: usize,
}

pub(crate) fn jupyter_nb_to_tokens(
    ipynb: &JupyterNotebook,
    start_rels_id: &mut usize,
    images: &mut BTreeMap<String, ImageFileContents>,
) -> Vec<Token> {
    let mut markdown_options = Options::empty();
    markdown_options.insert(Options::ENABLE_STRIKETHROUGH);

    let mut result: Vec<Token> = Vec::new();
    let mut image_counter: usize = 1;

    for cell in ipynb.cells.iter() {
        match cell.cell_type {
            JupyterCellType::Markdown => {
                let source_s = cell.source.join("\n");
                let tokens = markdown_to_tokens(&source_s);
                result.extend(tokens);
            }
            JupyterCellType::Code => {
                if let Some(outputs) = &cell.outputs {
                    for output in outputs.iter() {
                        if output.data.contains_key("text/plain") {
                            let value = output.data.get("text/plain").unwrap();
                            let text_lines: Vec<String> =
                                serde_json::from_value(value.clone()).unwrap();
                            for line in text_lines.iter() {
                                result.extend(monospace_paragraph_tokens(line));
                            }
                        }

                        if output.data.contains_key("image/png") {
                            let value = output.data.get("image/png").unwrap();
                            let base64_encoded_string: String =
                                serde_json::from_value(value.clone()).unwrap();

                            // N.B! Important to trim, because Jupyter seems to add an
                            // explicit newline character at the end of the Base64 string, for
                            // some reason.
                            if let Ok(payload) = base64::decode(base64_encoded_string.trim()) {
                                *start_rels_id += 1;
                                let figure_relationship_id = format!("rId{}", start_rels_id);
                                let filename = format!("figure-{}.png", image_counter);
                                let (width, height) = get_png_dimensions(&payload);
                                images.insert(
                                    figure_relationship_id.clone(),
                                    ImageFileContents {
                                        file_contents: FileContents { filename, payload },
                                        height: height,
                                        width: width,
                                    },
                                );

                                let tokens = image_paragraph_tokens(
                                    &figure_relationship_id,
                                    width,
                                    height,
                                    image_counter,
                                );

                                result.extend(tokens);

                                image_counter += 1;
                            } else {
                                // TODO do some proper error handling or error notifications
                                // here.
                            }
                        }
                    }
                }
            }
            JupyterCellType::Raw => {
                // Produce just raw text
            }
        }
    }

    result
}

fn get_png_dimensions(png_payload: &[u8]) -> (u32, u32) {
    let cursor = Cursor::new(png_payload);
    let decoder = png::Decoder::new(cursor);
    let (info, _) = decoder.read_info().unwrap();
    (info.width, info.height)
}
