mod accumulator;
mod compressor;

use crate::errors::TextkitDocxError;
use crate::{PageDimensions, TemplateArea, Token, TokenType, NS_WP_ML, PAT_HB_CPX, PAT_HB_SMP};
use accumulator::TemplateAccumulator;
use compressor::compress_tokens;
use regex::Regex;
use std::io::BufReader;
use std::io::{Read, Seek};
use xml::EventReader;
use zip::ZipArchive;

/// Reads a string of XML data and converts it into a vector
/// of Token objects.
pub(crate) fn xml_to_token_vec(xml: &str) -> Result<Vec<Token>, TextkitDocxError> {
    let mut result: Vec<Token> = Vec::new();

    let source_buf = BufReader::new(xml.as_bytes());
    let source_parser = EventReader::new(source_buf);

    let simple_template_pattern = Regex::new(PAT_HB_SMP).unwrap();
    let complex_template_pattern = Regex::new(PAT_HB_CPX).unwrap();
    let mut accumulator = TemplateAccumulator::Idle;

    for event in source_parser {
        match (&event, &accumulator) {
            (Ok(e @ xml::reader::XmlEvent::Characters(_)), _) => {
                if let xml::reader::XmlEvent::Characters(contents) = e {
                    accumulator.accumulate(contents);

                    if let TemplateAccumulator::Done(s) = &accumulator {
                        let token_type = if simple_template_pattern.is_match(&s) {
                            TokenType::Template
                        } else if complex_template_pattern.is_match(&s) {
                            TokenType::ComplexTemplate
                        } else {
                            TokenType::Normal
                        };

                        let new_event = xml::reader::XmlEvent::Characters(s.clone());

                        result.push(Token {
                            token_type: token_type,
                            token_text: Some(s.clone()),
                            xml_reader_event: new_event,
                        });

                        accumulator.reset();
                    }
                }
            }
            (Ok(anything_else), TemplateAccumulator::Idle) => result.push(Token {
                token_type: TokenType::Normal,
                token_text: None,
                xml_reader_event: anything_else.clone(),
            }),
            (Err(error), _) => return Err(error.clone().into()),
            _ => {}
        }
    }

    // FIXME here
    // compress_tokens(&result);

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

    // let debug_info: Vec<(usize, &Token)> = token_vec.iter().enumerate().collect();
    // println!("debu_info = {:#?}", debug_info);

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

pub(crate) fn unzip_text_file<T: Read + Seek>(
    archive: &mut ZipArchive<T>,
    file_name: &str,
) -> Result<String, TextkitDocxError> {
    let mut file = archive.by_name(file_name)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    Ok(contents)
}
