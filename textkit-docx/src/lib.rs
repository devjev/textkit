mod errors;

use crate::errors::TextkitDocxError;
use regex::Regex;
use serde::Serialize;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::{Path, PathBuf};
use xml::reader::EventReader;
use xml::writer::EmitterConfig;
use zip::{write::FileOptions, ZipArchive, ZipWriter};

/// Namespace string used in DOCX XML data to denote word processing elements (like paragraphs).
static NS_WP_ML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

type DocxPayload = ZipArchive<Cursor<Vec<u8>>>;

pub(crate) fn unzip_text_file<T: Read + Seek>(
    archive: &mut ZipArchive<T>,
    file_name: &str,
) -> Result<String, TextkitDocxError> {
    let mut file = archive.by_name(file_name)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    Ok(contents)
}

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

/// Reads a string of XML data and converts it into a vector
/// of Token objects.
pub(crate) fn xml_to_token_vec(xml: &str) -> Result<Vec<Token>, TextkitDocxError> {
    let mut result: Vec<Token> = Vec::new();

    let source_buf = BufReader::new(xml.as_bytes());
    let source_parser = EventReader::new(source_buf);

    let token_pattern = Regex::new("\\{\\{\\S+\\}\\}").unwrap();

    for event in source_parser {
        match &event {
            Ok(e @ xml::reader::XmlEvent::Characters(_)) => {
                if let xml::reader::XmlEvent::Characters(contents) = e {
                    if token_pattern.is_match(contents) {
                        result.push(Token {
                            token_type: TokenType::Template,
                            token_text: Some(contents.clone()),
                            xml_reader_event: e.clone(),
                        });
                    } else {
                        result.push(Token {
                            token_type: TokenType::Normal,
                            token_text: Some(contents.clone()),
                            xml_reader_event: e.clone(),
                        })
                    }
                }
            }
            Ok(anything_else) => result.push(Token {
                token_type: TokenType::Normal,
                token_text: None,
                xml_reader_event: anything_else.clone(),
            }),
            Err(error) => return Err(error.clone().into()),
        }
    }

    Ok(result)
}

pub(crate) fn write_token_vector_to_string_(
    tokens: &Vec<Token>,
) -> Result<String, TextkitDocxError> {
    let mut buf: Vec<u8> = Vec::new();
    let cursor = Cursor::new(&mut buf);
    let mut writer = EmitterConfig::new()
        .perform_indent(true)
        .create_writer(cursor);

    for item in tokens.iter() {
        if let Some(writer_event) = item.xml_reader_event.as_writer_event() {
            // the .write method returns a result, the error value of which is
            // of type xml::writer::emitter::EmitterError, which is private...
            // So here we are just passing along a token TextkitDocxError
            // instead.
            if let Err(_) = writer.write(writer_event) {
                return Err(TextkitDocxError::FailedWriteXml);
            }
        }
    }

    let result = String::from(std::str::from_utf8(&buf).unwrap());
    Ok(result)
}

pub(crate) fn new_zip_bytes_with_document_xml(
    zip_payload: &mut ZipArchive<Cursor<Vec<u8>>>,
    document_xml: &str,
) -> Result<Vec<u8>, TextkitDocxError> {
    // Prepare everything necessary to create a new zip payload
    // in memory.
    let mut buf: Vec<u8> = Vec::new();

    {
        let mut cursor = Cursor::new(&mut buf);
        let mut zip = ZipWriter::new(&mut cursor);
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o755);

        // We don't copy word/document.xml over.
        let excluded_path = Path::new("word/document.xml");

        // Copy over everything except for word/document.xml
        for i in 0..zip_payload.len() {
            // Extract the current file
            let mut file = zip_payload.by_index(i)?;

            // Write it to the new zip file
            if let Some(full_file_name) = file.sanitized_name().to_str() {
                let target_path = Path::new(full_file_name);

                if target_path != excluded_path {
                    let mut file_buf: Vec<u8> = Vec::new();
                    file.read_to_end(&mut file_buf)?;
                    zip.start_file_from_path(&target_path, options.clone())?;
                    zip.write_all(&file_buf)?;
                }
            }
        }

        zip.start_file_from_path(&excluded_path, options.clone())?;
        zip.write_all(document_xml.as_bytes())?;
        zip.finish()?;
    }

    Ok(buf)
}

pub(crate) fn find_template_areas(
    token_vec: &Vec<Token>,
    wrapping_element_name: &str,
) -> Vec<TemplateArea> {
    let mut result: Vec<TemplateArea> = Vec::new();
    let ns = Some(String::from(NS_WP_ML));

    let token_indices = token_vec.iter().enumerate().filter_map(|(i, token)| {
        if token.token_type == TokenType::Template {
            Some(i)
        } else {
            None
        }
    });

    for token_index in token_indices {
        let mut start_index: Option<usize> = None;
        let mut end_index: Option<usize> = None;

        {
            let mut anchor = token_index.clone();
            loop {
                match &token_vec[anchor].xml_reader_event {
                    xml::reader::XmlEvent::StartElement { name, .. } => {
                        if name.local_name == wrapping_element_name && name.namespace == ns {
                            start_index = Some(anchor);
                            break;
                        }
                    }
                    _ => (),
                }

                if anchor > 0 {
                    anchor = anchor - 1;
                } else {
                    break;
                }
            }
        } // find start_index

        {
            let mut anchor = token_index.clone();
            loop {
                match &token_vec[anchor].xml_reader_event {
                    xml::reader::XmlEvent::EndElement { name, .. } => {
                        if name.local_name == wrapping_element_name && name.namespace == ns {
                            end_index = Some(anchor);
                            break;
                        }
                    }
                    _ => (),
                }

                if anchor < token_vec.len() {
                    anchor = anchor + 1;
                } else {
                    break;
                }
            }
        } // find end_index

        result.push(TemplateArea {
            token_index,
            context_start_index: start_index,
            context_end_index: end_index,
        });
    }

    result
}

/// A .docx template supporting Handlebars syntax.
#[derive(Debug)]
pub struct DocxTemplate {
    pub source_payload: DocxPayload,
    document_xml: String,
}

impl DocxTemplate {
    /// Create DocxTemplate from memory.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, TextkitDocxError> {
        let buf = Vec::from(bytes);
        let cursor = Cursor::new(buf);
        let mut source_payload = ZipArchive::new(cursor)?;
        let document_xml = unzip_text_file(&mut source_payload, "word/document.xml")?;
        Ok(Self {
            source_payload,
            document_xml,
        })
    }

    /// Create a DocxTemplate from a .docx file on disk.
    pub fn from_file(file_name: &PathBuf) -> Result<Self, TextkitDocxError> {
        let mut fh = File::open(file_name)?;
        let mut buf: Vec<u8> = Vec::new();
        fh.read_to_end(&mut buf)?;
        DocxTemplate::from_bytes(&buf)
    }

    /// Render the template given some data context into a new .docx file (returned)
    /// as a vector of bytes.
    pub fn render<T: Serialize>(&self, data: &T) -> Result<Vec<u8>, TextkitDocxError> {
        let mut result: Vec<Token> = Vec::new();

        let tokens = xml_to_token_vec(&self.document_xml)?;
        let contexts = find_template_areas(&tokens, "p");
        let hb = handlebars::Handlebars::new();

        // This index tracks the position in the `tokens` vector of the last
        // non-template token that was processed.
        let mut bookmark_index: usize = 0;

        for context in contexts.iter() {
            if let TemplateArea {
                context_start_index: Some(start),
                context_end_index: Some(end),
                token_index: index,
            } = context
            {
                // The template area (expressed as a vector of tokens) identified
                // by the running TemplateArea.
                let subvector_index = index - start;
                let template_tokens = tokens[*start..=*end].to_vec();

                // All non-template tokens between the last template area and
                // the current one.
                let prequel = tokens[bookmark_index..*start].to_vec();

                // Set the bookmark_index to after the end of the current template
                // area for the next iteration.
                bookmark_index = end + 1;

                // Fill the result with non-template tokens preceeding this template.
                result.extend(prequel);

                // Process the template.
                if let Some(template_text) = &template_tokens[subvector_index].token_text {
                    match hb.render_template(template_text, data) {
                        Ok(rendered_text) if !rendered_text.is_empty() => {
                            let rendered_chunks = rendered_text.split("\n\n");

                            for chunk in rendered_chunks {
                                let chunk_text = String::from(chunk);
                                let mut chunk_tokens = template_tokens.clone();
                                chunk_tokens[subvector_index] = Token {
                                    token_type: TokenType::Normal,
                                    token_text: Some(chunk_text.clone()),
                                    xml_reader_event: xml::reader::XmlEvent::Characters(
                                        chunk_text.clone(),
                                    ),
                                };
                                result.extend(chunk_tokens);
                            }
                        }
                        _ => {
                            result.extend(template_tokens);
                        }
                    }
                }
            }
        }

        // Add remaining tokens to the result
        let sequel = tokens[bookmark_index..].to_vec();
        result.extend(sequel);

        // New document.xml contents
        let document_xml_contents = write_token_vector_to_string_(&result)?;

        // NOTE Not sure if cloning here is really necessary.
        let mut payload = self.source_payload.clone();

        new_zip_bytes_with_document_xml(&mut payload, &document_xml_contents)
    }
}
