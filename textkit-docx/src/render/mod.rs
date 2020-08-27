pub mod datakit_table;
pub mod jupyter_nb;

use crate::errors::TextkitDocxError;
use crate::{Token, TokenType, NS_WP_ML};
use serde::Serialize;
use std::collections::BTreeMap;
use std::io::Cursor;
use std::io::{Read, Write};
use std::path::Path;
use xml::writer::EmitterConfig;
use zip::{write::FileOptions, ZipArchive, ZipWriter};

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

pub(crate) fn owned_name(prefix: &str, tag_name: &str) -> xml::name::OwnedName {
    xml::name::OwnedName {
        local_name: tag_name.into(),
        namespace: Some(NS_WP_ML.into()),
        prefix: Some(prefix.into()),
    }
}

pub(crate) fn owned_attributes(
    attrs: &[(&str, &str, &str)],
) -> Vec<xml::attribute::OwnedAttribute> {
    let mut result: Vec<xml::attribute::OwnedAttribute> = Vec::new();
    for attr in attrs.iter() {
        let (prefix, name, value) = attr;
        let owned_attribute = xml::attribute::OwnedAttribute {
            name: owned_name(prefix, name),
            value: String::from(*value),
        };
        result.push(owned_attribute);
    }
    result
}

pub(crate) fn start_tag_event(
    tag_name: &str,
    attrs: Option<&[(&str, &str, &str)]>,
) -> xml::reader::XmlEvent {
    let mut ns: BTreeMap<String, String> = BTreeMap::new();
    ns.insert("w".into(), NS_WP_ML.into());

    let attributes: Vec<xml::attribute::OwnedAttribute> = if let Some(supplied_attrs) = attrs {
        owned_attributes(supplied_attrs)
    } else {
        vec![]
    };

    xml::reader::XmlEvent::StartElement {
        name: owned_name("w", tag_name),
        namespace: xml::namespace::Namespace(ns),
        attributes: attributes,
    }
}

pub(crate) fn end_tag_event(tag_name: &str) -> xml::reader::XmlEvent {
    xml::reader::XmlEvent::EndElement {
        name: owned_name("w", tag_name),
    }
}

pub(crate) fn paragraph_tokens(contents: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let paragraphs = split_string_by_empty_line(contents);

    for paragraph in paragraphs {
        let prequel = paragraph_prequel_tokens();
        let run_start = run_start_token();
        let run_end = run_end_token();
        let chars = char_text_tokens(paragraph, true);
        let sequel = paragraph_sequel_tokens();

        result.extend(prequel);
        result.push(run_start);
        result.extend(chars);
        result.push(run_end);
        result.extend(sequel);
    }

    result
}

pub(crate) fn monospace_paragraph_tokens(contents: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let paragraphs = split_string_by_empty_line(contents);

    for paragraph in paragraphs {
        let prequel = paragraph_prequel_tokens();
        let run_start = run_start_token();
        let run_end = run_end_token();
        let chars = char_text_tokens(paragraph, true);
        let sequel = paragraph_sequel_tokens();

        result.extend(prequel);

        /* add this to paragraph prequel
        <w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
                <w:sz w:val="16"/>
                <w:szCs w:val="16"/>
            </w:rPr>
        </w:pPr>
        */
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("pPr", None),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("rPr", None),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event(
                "rFonts",
                Some(&[("w", "ascii", "Consolas"), ("w", "hAnsi", "Consolas")]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("rFonts"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("sz", Some(&[("w", "val", "16")])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("sz"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("szCs", Some(&[("w", "val", "16")])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("szCs"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("rPr"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("pPr"),
            token_text: None,
        });

        result.push(run_start);

        /* add this to run prequel
        <w:rPr>
            <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
            <w:sz w:val="16"/>
            <w:szCs w:val="16"/>
        </w:rPr>
        */
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("rPr", None),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event(
                "rFonts",
                Some(&[("w", "ascii", "Consolas"), ("w", "hAnsi", "Consolas")]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("rFonts"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("sz", Some(&[("w", "val", "16")])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("sz"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event("szCs", Some(&[("w", "val", "16")])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("szCs"),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event("rPr"),
            token_text: None,
        });
        result.extend(chars);
        result.push(run_end);
        result.extend(sequel);
    }

    result
}

pub(crate) fn paragraph_prequel_tokens() -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("p", None),
        token_text: None,
    });

    result
}

pub(crate) fn heading_prequel_tokens(heading_style: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("p", None),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("pPr", None),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("pStyle", Some(&[("w", "val", heading_style)])),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event("pStyle"),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event("pPr"),
        token_text: None,
    });

    result
}

pub(crate) fn paragraph_sequel_tokens() -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    // </w:p>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event("p"),
        token_text: None,
    });

    result
}

pub(crate) fn heading_sequel_tokens() -> Vec<Token> {
    paragraph_sequel_tokens()
}

pub(crate) fn char_text_tokens(contents: &str, preserve_space: bool) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let attrs: Option<&[(&str, &str, &str)]> = if preserve_space {
        Some(&[("xml", "space", "preserve")])
    } else {
        None
    };

    // <w:t>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("t", attrs),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: xml::reader::XmlEvent::Characters(contents.into()),
        token_text: Some(contents.into()),
    });

    // </w:t>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event("t"),
        token_text: None,
    });

    result
}

pub(crate) fn run_start_token() -> Token {
    Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event("r", None),
        token_text: None,
    }
}

pub(crate) fn run_end_token() -> Token {
    Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event("r"),
        token_text: None,
    }
}

// pub(crate) fn heading_tokens(contents: &str, heading_style: &str) -> Vec<Token> {
//     let mut result: Vec<Token> = Vec::new();
//     let paragraphs = split_string_by_empty_line(contents);

//     for paragraph in paragraphs {
//         let prequel = heading_prequel_tokens(heading_style);
//         let run_start = run_start_token();
//         let run_end = run_end_token();
//         let chars = char_text_tokens(paragraph, true);
//         let sequel = heading_sequel_tokens();
//         result.extend(prequel);
//         result.push(run_start);
//         result.extend(chars);
//         result.push(run_end);
//         result.extend(sequel);
//     }

//     result
// }

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
                // to the XML code produced by `paragraph_tokens(contents: &str)`,
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
