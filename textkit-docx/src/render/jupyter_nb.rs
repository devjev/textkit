//! Render Jupyter Notebooks
//!

use crate::render::{make_end_tag_event, make_paragraph_tokens, make_start_tag_event};
use crate::{Token, TokenType};
use pulldown_cmark::{Options, Parser};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

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

pub fn jupyter_nb_to_tokens_proto(ipynb: &JupyterNotebook) {
    let mut markdown_options = Options::empty();
    markdown_options.insert(Options::ENABLE_STRIKETHROUGH);

    let mut result: Vec<Token> = Vec::new();

    for cell in ipynb.cells.iter() {
        match cell.cell_type {
            JupyterCellType::Markdown => {
                let source_s = cell.source.join("\n");
                let parser = Parser::new_ext(&source_s, markdown_options);
                for parser_event in parser {
                    match parser_event {
                        pulldown_cmark::Event::Start(tag) => cmark_tag_to_wp_tag_start(&tag),
                        pulldown_cmark::Event::Text(x) => println!("text {:#?}", x),
                        _ => println!("{:#?}", parser_event),
                    }
                }
            }
            JupyterCellType::Code => {
                if let Some(outputs) = &cell.outputs {
                    // Get either an HTML output or
                    // Get image/png
                    //println!("{:#?}", outputs);
                }
            }
            JupyterCellType::Raw => {
                // Produce just raw text
            }
        }
    }
}

pub(crate) fn jupyter_nb_to_tokens(ipynb: &JupyterNotebook) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    for cell in ipynb.cells.iter() {
        match cell.cell_type {
            JupyterCellType::Markdown => {
                let source_s = cell.source.join("\n");
                // parse markdown here
            }
            JupyterCellType::Code => {
                if let Some(outputs) = &cell.outputs {
                    // Get either an HTML output or
                    // Get image/png
                    println!("{:#?}", outputs);
                }
            }
            JupyterCellType::Raw => {
                // Produce just raw text
            }
        }
    }
    result
}

fn cmark_tag_to_wp_tag_start(cmark_tag: &pulldown_cmark::Tag) {
    //make_start_tag_event(tag_name: &str, attrs: Option<&[(&str, &str)]>)
}
