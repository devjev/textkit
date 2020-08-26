pub mod errors;
pub mod parse;
pub mod render;
pub mod template;

pub use crate::template::DocxTemplate;

use std::io::Cursor;
use zip::ZipArchive;

/// Namespace string used in DOCX XML data to denote word processing elements (like paragraphs).
static NS_WP_ML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Regex patterns used to match Handlebars placeholders
static PAT_HB_ALL: &str = r"\{\{(\S+)\s*([^\{\}]+)?\}\}"; // All placeholders
static PAT_HB_SMP: &str = r"\{\{\S+\}\}"; // Only simple placeholders
static PAT_HB_CPX: &str = r"\{\{[^#/](\S+)\s+([^\{\}]+)?\}\}"; // Only placeholders with helpers
static PAT_HB_MLS: &str = r"\{\{#(.+)\}\}";
static PAT_HB_MLE: &str = r"\{\{\\/(.+)\}\}";

type DocxPayload = ZipArchive<Cursor<Vec<u8>>>;

/// `textkit-docx` treats XML data as a vector of tokens, which can
/// represent a opening tag, a closing tag, CDATA, character data, etc.
#[derive(Debug, Clone)]
pub(crate) struct Token {
    token_type: TokenType,
    token_text: Option<String>,
    xml_reader_event: xml::reader::XmlEvent,
}

/// Differentiates between tokens that contain character data with
/// Handlebars templating syntax and everything else.
#[derive(Debug, PartialEq, Clone)]
pub(crate) enum TokenType {
    Template,
    ComplexTemplate,
    Normal,
}

/// The indices indicating the first and last tokens around
/// a template placeholder.
#[derive(Debug)]
pub(crate) struct TemplateArea {
    pub context_start_index: Option<usize>,
    pub context_end_index: Option<usize>,
    pub token_index: usize,
}

#[derive(Debug)]
pub(crate) struct TemplatePlaceholder {
    pub helper_name: Option<String>,
    pub expression: String,
    pub start_position: usize,
    pub end_position: usize,
}

#[derive(Debug)]
pub struct PageDimensions {
    pub height: i32,
    pub width: i32,
    pub m_top: i32,
    pub m_bottom: i32,
    pub m_right: i32,
    pub m_left: i32,
    pub header: i32,
    pub footer: i32,
    pub gutter: i32,
}
