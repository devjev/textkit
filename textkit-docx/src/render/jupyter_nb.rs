//! Render Jupyter Notebooks
//!

use crate::render::{
    char_text_tokens, end_tag_event, heading_prequel_tokens, heading_sequel_tokens,
    monospace_paragraph_tokens, paragraph_prequel_tokens, paragraph_sequel_tokens, run_end_token,
    run_start_token, start_tag_event,
};
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

pub(crate) fn jupyter_nb_to_tokens(ipynb: &JupyterNotebook) -> Vec<Token> {
    let mut markdown_options = Options::empty();
    markdown_options.insert(Options::ENABLE_STRIKETHROUGH);

    let mut result: Vec<Token> = Vec::new();

    for cell in ipynb.cells.iter() {
        match cell.cell_type {
            JupyterCellType::Markdown => {
                let source_s = cell.source.join("\n");
                let parser = Parser::new_ext(&source_s, markdown_options);
                let mut is_inline = false;

                for parser_event in parser {
                    match parser_event {
                        pulldown_cmark::Event::Start(tag) => {
                            result.extend(cmark_tag_to_wp_tag_start(&tag, &mut is_inline));
                        }
                        pulldown_cmark::Event::Text(x) => {
                            if is_inline {
                                let text_with_spaces = format!(" {} ", x);
                                result.extend(char_text_tokens(&text_with_spaces, false));
                            } else {
                                result.push(run_start_token());
                                result.extend(char_text_tokens(&x, true));
                                result.push(run_end_token());
                            }
                        }
                        pulldown_cmark::Event::End(tag) => {
                            result.extend(cmark_tag_to_wp_tag_end(&tag, &mut is_inline));
                        }
                        _ => {}
                    }
                }
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

                        // TODO deal with "image/png" here
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

// Goal here is to produce a vector of tokens for each cmark_tag.
fn cmark_tag_to_wp_tag_start(cmark_tag: &pulldown_cmark::Tag, is_inline: &mut bool) -> Vec<Token> {
    match cmark_tag {
        pulldown_cmark::Tag::Heading(level) => {
            let heading_style = format!("Heading{}", level + 1);
            heading_prequel_tokens(&heading_style)
        }
        pulldown_cmark::Tag::Paragraph => paragraph_prequel_tokens(),
        pulldown_cmark::Tag::Emphasis => {
            *is_inline = true;
            let mut emphasis_start: Vec<Token> = Vec::new();
            emphasis_start.push(run_start_token());
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: start_tag_event("rPr", None),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: start_tag_event("i", None),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: end_tag_event("i"),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: start_tag_event("iCs", None),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: end_tag_event("iCs"),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: end_tag_event("rPr"),
                token_text: None,
            });
            emphasis_start
        }
        _ => vec![],
    }
}

fn cmark_tag_to_wp_tag_end(cmark_tag: &pulldown_cmark::Tag, is_inline: &mut bool) -> Vec<Token> {
    match cmark_tag {
        pulldown_cmark::Tag::Heading(_) => heading_sequel_tokens(),
        pulldown_cmark::Tag::Paragraph => paragraph_sequel_tokens(),
        pulldown_cmark::Tag::Emphasis => {
            *is_inline = false;
            vec![run_end_token()]
        }
        _ => vec![],
    }
}
