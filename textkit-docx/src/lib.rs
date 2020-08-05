mod errors;

use crate::errors::TextkitDocxError;
use datakit::table::Table;
use datakit::value::definitions::*;
use datakit::value::primitives::*;
use regex::Regex;
use serde::Serialize;
use std::collections::{BTreeMap, HashSet};
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::{Path, PathBuf};
use xml::reader::EventReader;
use xml::writer::EmitterConfig;
use zip::{write::FileOptions, ZipArchive, ZipWriter};

/// Namespace string used in DOCX XML data to denote word processing elements (like paragraphs).
static NS_WP_ML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Regex patterns used to match Handlebars placeholders
static PAT_HB_ALL: &str = "\\{\\{(\\S+)\\s*([^\\{\\}]+)?\\}\\}"; // All placeholders
static PAT_HB_SMP: &str = "\\{\\{\\S+\\}\\}"; // Only simple placeholders
static PAT_HB_CPX: &str = "\\{\\{(\\S+)\\s+([^\\{\\}]+)?\\}\\}"; // Only placeholders with helpers

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

/// A .docx template supporting Handlebars syntax.
#[derive(Debug)]
pub struct DocxTemplate {
    source_payload: DocxPayload,
    document_xml: String,
    tokens: Vec<Token>,
    dimensions: PageDimensions,
    template_areas: Vec<TemplateArea>,
}

impl DocxTemplate {
    /// Create DocxTemplate from memory.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, TextkitDocxError> {
        let buf = Vec::from(bytes);
        let cursor = Cursor::new(buf);

        let mut source_payload = ZipArchive::new(cursor)?;
        let document_xml = unzip_text_file(&mut source_payload, "word/document.xml")?;

        let tokens = xml_to_token_vec(&document_xml)?;
        let dimensions = parse_page_dimensions(&document_xml)?;
        let template_areas = find_template_areas(&tokens, "p");

        Ok(Self {
            source_payload,
            document_xml,
            tokens,
            dimensions,
            template_areas,
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

        // This index tracks the position in the `self.tokens` vector of the last
        // non-template token that was processed.
        let mut bookmark_index: usize = 0;

        // We need to keep track of all template area start indices, i.e. where
        // the paragraph that contains template placeholders starts. If there
        // are multiple placeholders in a single paragraph, we want to render
        // that paragraph only **once**, and not repeat for every placeholder.
        // For that we will keep track of all the seen indices, and if a
        // two or more different TemplateAreas have the same start index, the
        // rendering happens only at the first TemplateArea, so as to avoid
        // duplicates.
        let mut already_seen_start_indices: HashSet<usize> = HashSet::new();

        // Also, we need a json serialized version of the data (mimicing Handlebars)
        // to render custom complex templates.
        let serialized_data = serde_json::to_value(data)?;

        for context in self.template_areas.iter() {
            if let TemplateArea {
                context_start_index: Some(start),
                context_end_index: Some(end),
                token_index: index,
            } = context
            {
                // If the another TemplateArea already was in the same
                // start/end range, skip it. Here we assume that for each
                // start index there is always the same end index.
                if already_seen_start_indices.contains(start) {
                    continue;
                } else {
                    already_seen_start_indices.insert(*start);
                }

                // The template area (expressed as a vector of tokens) identified
                // by the running TemplateArea.
                let subvector_index = index - start;
                let template_tokens = self.tokens[*start..=*end].to_vec();

                // All non-template tokens between the last template area and
                // the current one.
                let prequel = self.tokens[bookmark_index..*start].to_vec();

                // Set the bookmark_index to after the end of the current template
                // area for the next iteration.
                bookmark_index = end + 1;

                // Fill the result with non-template tokens preceeding this template.
                result.extend(prequel);

                // Process the template.
                if let Some(template_text) = &template_tokens[subvector_index].token_text {
                    match &template_tokens[subvector_index].token_type {
                        TokenType::Template => {
                            let tokens = render_and_paste_tokens(
                                &template_tokens,
                                template_text,
                                subvector_index,
                                data,
                            );
                            result.extend(tokens);
                        }
                        TokenType::ComplexTemplate => {
                            let mut index: usize = 0;
                            let placeholders = parse_template_placeholders(template_text);

                            for placeholder in placeholders.iter() {
                                if let Some(helper_name) = &placeholder.helper_name {
                                    if index != placeholder.start_position {
                                        let prequel =
                                            &template_text[index..placeholder.start_position];

                                        let prequel_tokens = render_and_paste_tokens(
                                            &template_tokens,
                                            prequel,
                                            subvector_index,
                                            data,
                                        );
                                        result.extend(prequel_tokens);
                                    }

                                    index = placeholder.end_position;

                                    // Right now, there is only one complex template helper:
                                    // `table` that generates a table if provided with a datakit::Table.
                                    // Future versions should include dynamic handling of this.
                                    if helper_name == "table" {
                                        // get the bit that's datakit table
                                        if let Some(table_serialized) =
                                            serialized_data.get(&placeholder.expression)
                                        {
                                            let table: Table =
                                                serde_json::from_value(table_serialized.clone())?;
                                            let table_tokens =
                                                datakit_table_to_tokens(&table, &self.dimensions);
                                            result.extend(table_tokens);
                                        }
                                    }
                                }
                            }
                            if index != template_text.len() {
                                let sequel = &template_text[index..];
                                let sequel_tokens = render_and_paste_tokens(
                                    &template_tokens,
                                    sequel,
                                    subvector_index,
                                    data,
                                );
                                result.extend(sequel_tokens);
                            }
                        }
                        _ => (),
                    }
                }
            }
        }

        // Add remaining tokens to the result
        let sequel = self.tokens[bookmark_index..].to_vec();
        result.extend(sequel);

        // New document.xml contents
        let document_xml_contents = write_token_vector_to_string(&result)?;

        // NOTE Not sure if cloning here is really necessary.
        let mut payload = self.source_payload.clone();

        new_zip_bytes_with_document_xml(&mut payload, &document_xml_contents)
    }
}

pub(crate) fn parse_template_placeholders(text: &str) -> Vec<TemplatePlaceholder> {
    let mut result: Vec<TemplatePlaceholder> = Vec::new();
    let placeholder_pattern = Regex::new(PAT_HB_ALL).unwrap();

    for capture in placeholder_pattern.captures_iter(text) {
        let start_position = text.find(&capture[0]).unwrap();
        let end_position = start_position + &capture[0].len();
        if capture.len() == 3 {
            result.push(TemplatePlaceholder {
                helper_name: Some(capture[1].into()),
                expression: capture[2].into(),
                start_position: start_position,
                end_position: end_position,
            })
        } else {
            result.push(TemplatePlaceholder {
                helper_name: None,
                expression: capture[1].into(),
                start_position: start_position,
                end_position: end_position,
            })
        }
    }

    result
}

/// Reads a string of XML data and converts it into a vector
/// of Token objects.
pub(crate) fn xml_to_token_vec(xml: &str) -> Result<Vec<Token>, TextkitDocxError> {
    let mut result: Vec<Token> = Vec::new();

    let source_buf = BufReader::new(xml.as_bytes());
    let source_parser = EventReader::new(source_buf);

    let simple_template_pattern = Regex::new(PAT_HB_SMP).unwrap();
    let complex_template_pattern = Regex::new(PAT_HB_CPX).unwrap();

    for event in source_parser {
        match &event {
            Ok(e @ xml::reader::XmlEvent::Characters(_)) => {
                if let xml::reader::XmlEvent::Characters(contents) = e {
                    if simple_template_pattern.is_match(contents) {
                        result.push(Token {
                            token_type: TokenType::Template,
                            token_text: Some(contents.clone()),
                            xml_reader_event: e.clone(),
                        });
                    } else if complex_template_pattern.is_match(contents) {
                        result.push(Token {
                            token_type: TokenType::ComplexTemplate,
                            token_text: Some(contents.clone()),
                            xml_reader_event: e.clone(),
                        })
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

/// Extract page dimensions from DOCX data.
pub(crate) fn parse_page_dimensions(
    document_xml: &str,
) -> Result<PageDimensions, TextkitDocxError> {
    let mut height_opt: Option<i32> = None;
    let mut width_opt: Option<i32> = None;
    let mut m_top_opt: Option<i32> = None;
    let mut m_bottom_opt: Option<i32> = None;
    let mut m_right_opt: Option<i32> = None;
    let mut m_left_opt: Option<i32> = None;
    let mut header_opt: Option<i32> = None;
    let mut footer_opt: Option<i32> = None;
    let mut gutter_opt: Option<i32> = None;

    let source_buf = BufReader::new(document_xml.as_bytes());
    let parser = EventReader::new(source_buf);
    let ns = Some(String::from(NS_WP_ML));

    let fetch_attr_value = |attrs: &[xml::attribute::OwnedAttribute], name: &str| {
        for attr in attrs.iter() {
            if attr.name.local_name == name && attr.name.namespace == ns {
                return Some(attr.value.parse::<i32>().unwrap());
            }
        }
        None
    };

    for event in parser {
        match &event {
            Ok(xml::reader::XmlEvent::StartElement {
                name, attributes, ..
            }) => {
                // Fetch page size
                if name.local_name == "pgSz" && name.namespace == ns {
                    width_opt = fetch_attr_value(attributes, "w");
                    height_opt = fetch_attr_value(attributes, "h");
                }

                // Fetch page margins
                if name.local_name == "pgMar" && name.namespace == ns {
                    m_top_opt = fetch_attr_value(attributes, "top");
                    m_bottom_opt = fetch_attr_value(attributes, "bottom");
                    m_right_opt = fetch_attr_value(attributes, "right");
                    m_left_opt = fetch_attr_value(attributes, "left");
                    header_opt = fetch_attr_value(attributes, "header");
                    footer_opt = fetch_attr_value(attributes, "footer");
                    gutter_opt = fetch_attr_value(attributes, "gutter");
                }
            }
            _ => (),
        }
    }

    if height_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page height attribute not found".into(),
        ))
    } else if width_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page width attribute not found".into(),
        ))
    } else if m_top_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page top margin attribute not found".into(),
        ))
    } else if m_bottom_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page bottom margin attribute not found".into(),
        ))
    } else if m_right_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page right margin attribute not found".into(),
        ))
    } else if m_left_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page left margin attribute not found".into(),
        ))
    } else if header_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page header margin attribute not found".into(),
        ))
    } else if footer_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page footer margin attribute not found".into(),
        ))
    } else if gutter_opt == None {
        Err(TextkitDocxError::Malformed(
            "Page gutter attribute not found".into(),
        ))
    } else {
        Ok(PageDimensions {
            height: height_opt.unwrap(),
            width: width_opt.unwrap(),
            m_top: m_top_opt.unwrap(),
            m_bottom: m_bottom_opt.unwrap(),
            m_right: m_right_opt.unwrap(),
            m_left: m_left_opt.unwrap(),
            header: header_opt.unwrap(),
            footer: footer_opt.unwrap(),
            gutter: gutter_opt.unwrap(),
        })
    }
}

pub(crate) fn find_template_areas(
    token_vec: &Vec<Token>,
    wrapping_element_name: &str,
) -> Vec<TemplateArea> {
    let mut result: Vec<TemplateArea> = Vec::new();
    let ns = Some(String::from(NS_WP_ML));

    let token_indices = token_vec.iter().enumerate().filter_map(|(i, token)| {
        if token.token_type == TokenType::Template || token.token_type == TokenType::ComplexTemplate
        {
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

pub(crate) fn write_token_vector_to_string(
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
            if let Err(error) = writer.write(writer_event) {
                println!("{:?}", error);
                return Err(TextkitDocxError::FailedWriteXml);
            }
        }
    }

    let result = String::from(std::str::from_utf8(&buf).unwrap());
    Ok(result)
}

pub(crate) fn unzip_text_file<T: Read + Seek>(
    archive: &mut ZipArchive<T>,
    file_name: &str,
) -> Result<String, TextkitDocxError> {
    let mut file = archive.by_name(file_name)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    Ok(contents)
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

pub(crate) fn split_string_by_empty_line(string: &str) -> std::str::Split<&str> {
    if string.contains("\r\n") {
        string.split("\r\n\r\n")
    } else {
        string.split("\n\n")
    }
}

pub(crate) fn make_owned_name(tag_name: &str) -> xml::name::OwnedName {
    xml::name::OwnedName {
        local_name: tag_name.into(),
        namespace: Some(NS_WP_ML.into()),
        prefix: Some("w".into()),
    }
}

pub(crate) fn make_owned_attributes(attrs: &[(&str, &str)]) -> Vec<xml::attribute::OwnedAttribute> {
    let mut result: Vec<xml::attribute::OwnedAttribute> = Vec::new();
    for attr in attrs.iter() {
        let (name, value) = attr;
        let owned_attribute = xml::attribute::OwnedAttribute {
            name: make_owned_name(name),
            value: String::from(*value),
        };
        result.push(owned_attribute);
    }
    result
}

pub(crate) fn make_start_tag_event(
    tag_name: &str,
    attrs: Option<&[(&str, &str)]>,
) -> xml::reader::XmlEvent {
    let mut ns: BTreeMap<String, String> = BTreeMap::new();
    ns.insert("w".into(), NS_WP_ML.into());

    let attributes: Vec<xml::attribute::OwnedAttribute> = if let Some(supplied_attrs) = attrs {
        make_owned_attributes(supplied_attrs)
    } else {
        vec![]
    };

    xml::reader::XmlEvent::StartElement {
        name: make_owned_name(tag_name),
        namespace: xml::namespace::Namespace(ns),
        attributes: attributes,
    }
}

pub(crate) fn make_end_tag_event(tag_name: &str) -> xml::reader::XmlEvent {
    xml::reader::XmlEvent::EndElement {
        name: make_owned_name(tag_name),
    }
}

pub(crate) fn make_paragraph_tokens(contents: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let paragraphs = split_string_by_empty_line(contents);

    for paragraph in paragraphs {
        // <w:p>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_start_tag_event("p", None),
            token_text: None,
        });

        // <w:r>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_start_tag_event("r", None),
            token_text: None,
        });

        // <w:t>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_start_tag_event("t", None),
            token_text: None,
        });

        // Put character data inside the <w:t> element
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: xml::reader::XmlEvent::Characters(paragraph.into()),
            token_text: Some(paragraph.into()),
        });

        // </w:t>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_end_tag_event("t"),
            token_text: None,
        });

        // </w:t>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_end_tag_event("r"),
            token_text: None,
        });

        // </w:t>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_end_tag_event("p"),
            token_text: None,
        });
    }

    result
}

pub(crate) fn render_and_paste_tokens<T: Serialize>(
    template_tokens: &[Token],
    template_text: &str,
    token_index_to_replace: usize,
    data: &T,
) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let hb = handlebars::Handlebars::new();

    match hb.render_template(template_text, data) {
        Ok(rendered_text) if !rendered_text.is_empty() => {
            // Here for each paragraph in the rendered text, we take the
            // paragraph formating (and other attributes coming with it)
            // of where the placeholder was located and use it to produce
            // DOCX paragraphs in the template. Note, that this is different
            // to the XML code produced by `make_paragraph_tokens(contents: &str)`,
            // since that function does not retain any formatting and just
            // produces "naked" paragraphs.
            let rendered_chunks = split_string_by_empty_line(&rendered_text);
            for chunk in rendered_chunks {
                let chunk_text = String::from(chunk);

                // Copy the template tokens
                let mut chunk_tokens: Vec<Token> = template_tokens.clone().into();

                // ... and replace the part containing the placeholder
                // with our rendered content.
                chunk_tokens[token_index_to_replace] = Token {
                    token_type: TokenType::Normal,
                    token_text: Some(chunk_text.clone()),
                    xml_reader_event: xml::reader::XmlEvent::Characters(chunk_text.clone()),
                };
                result.extend(chunk_tokens);
            }
        }
        _ => {
            let tokens: Vec<Token> = template_tokens.clone().into();
            result.extend(tokens);
        }
    }

    result
}

pub(crate) fn datakit_table_to_tokens(table: &Table, dims: &PageDimensions) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    // <w:tbl>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event("tbl", None),
        token_text: None,
    });

    // <w:tblPr>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event("tblPr", None),
        token_text: None,
    });

    // <w:tblStyle w:val="TableGrid" />
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event("tblStyle", Some(&[("val", "TableGrid")])),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tblStyle"),
        token_text: None,
    });

    // <w:tblW w:w="0" w:type="auto" />
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event("tblW", Some(&[("w", "0"), ("type", "auto")])),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tblW"),
        token_text: None,
    });

    // <w:tblLook
    //     w:val="04A0"
    //     w:firstRow="1"
    //     w:lastRow="0"
    //     w:firstColumn="1"
    //     w:lastColumn="0"
    //     w:noHBand="0"
    //     w:noVBand="1" />
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event(
            "tblLook",
            Some(&[
                ("val", "04A0"),
                ("firstRow", "1"),
                ("lastRow", "0"),
                ("firstColumn", "1"),
                ("lastColumn", "0"),
                ("noHBand", "0"),
                ("noVBand", "1"),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tblLook"),
        token_text: None,
    });

    // </w:tblPr>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tblPr"),
        token_text: None,
    });

    // <w:tblGrid>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_start_tag_event("tblGrid", None),
        token_text: None,
    });

    // Now we need to figure out how many columns there are and what
    // the individual column width is (assuming equal widths).
    let no_of_cols = table.columns().len();
    let col_width = (dims.width - dims.m_left - dims.m_right) / 2 + 15;
    let col_width_str = format!("{}", col_width);

    for _ in 0..no_of_cols {
        // <w:gridCol w:w="<COL_WIDTH>"
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_start_tag_event("gridCol", Some(&[("w", &col_width_str)])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_end_tag_event("gridCol"),
            token_text: None,
        });
    }

    // <w:tblGrid>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tblGrid"),
        token_text: None,
    });

    // Now we need to populate the contents of the the table
    for row_i in 0..table.len() {
        // <w:tr> - we deliberately omitting any kind of id attributes (like w:rsidR).
        // That's not very compliant as far as I know, but MS Word handles it pretty
        // well, so IDs are a TODO for future versions.
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_start_tag_event("tr", None),
            token_text: None,
        });

        // Populate all table cells for the current row.
        for col_i in 0..no_of_cols {
            // <w:tc>
            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_start_tag_event("tc", None),
                token_text: None,
            });

            // <w:tcPr>
            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_start_tag_event("tcPr", None),
                token_text: None,
            });

            // <w:tcW w:w="<COL_WIDTH>" w:type="dxa" />
            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_start_tag_event(
                    "tcW",
                    Some(&[("w", &col_width_str), ("type", "dxa")]),
                ),
                token_text: None,
            });
            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_end_tag_event("tcW"),
                token_text: None,
            });

            // </w:tcPr>
            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_end_tag_event("tcPr"),
                token_text: None,
            });

            // TODO fill this thing with data from the value
            let cell_value = &table.columns()[col_i][row_i];

            match cell_value {
                Value::Text(text) => {
                    let paragraph_tokens = make_paragraph_tokens(text);
                    result.extend(paragraph_tokens);
                }
                Value::Number(Numeric::Integer(int)) => {
                    let int_str = format!("{}", int);
                    let paragraph_tokens = make_paragraph_tokens(&int_str);
                    result.extend(paragraph_tokens);
                }
                Value::Number(Numeric::Real(real)) => {
                    let real_str = format!("{:.3}", real);
                    let paragraph_tokens = make_paragraph_tokens(&real_str);
                    result.extend(paragraph_tokens);
                }
                Value::Boolean(boolean) => {
                    let bool_str = if *boolean { "Yes" } else { "No" };
                    let paragraph_tokens = make_paragraph_tokens(&bool_str);
                    result.extend(paragraph_tokens);
                }
                Value::DateTime(dt) => {
                    let dt_str = format!("{}", dt);
                    let paragraph_tokens = make_paragraph_tokens(&dt_str);
                    result.extend(paragraph_tokens);
                }
                // TODO implement the rest of value types
                _ => (),
            }

            result.push(Token {
                token_type: TokenType::Normal,
                xml_reader_event: make_end_tag_event("tc"),
                token_text: None,
            });
        }

        // </w:tr>
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: make_end_tag_event("tr"),
            token_text: None,
        });
    }

    // </w:tbl>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: make_end_tag_event("tbl"),
        token_text: None,
    });

    result
}
