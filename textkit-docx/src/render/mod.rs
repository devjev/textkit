use crate::errors::TextkitDocxError;
use crate::{PageDimensions, Token, TokenType, NS_WP_ML};
use datakit::{table::Table, value::definitions::*, value::primitives::*};
use serde::Serialize;
use std::collections::BTreeMap;
use std::io::Cursor;
use std::io::{Read, Write};
use std::path::Path;
use xml::writer::EmitterConfig;
use zip::{write::FileOptions, ZipArchive, ZipWriter};

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
        //Ok(rendered_text) if !rendered_text.is_empty() => {
        Ok(rendered_text) => {
            if !rendered_text.is_empty() {
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
        }
        _ => {
            let tokens: Vec<Token> = template_tokens.clone().into();
            result.extend(tokens);
        }
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
            if let Err(_) = writer.write(writer_event) {
                return Err(TextkitDocxError::FailedWriteXml);
            }
        }
    }

    let result = String::from(std::str::from_utf8(&buf).unwrap());
    Ok(result)
}
