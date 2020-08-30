pub mod datakit_table;
pub mod jupyter_nb;

use crate::errors::TextkitDocxError;
use crate::{
    ImageFileContents, Token, TokenType, NS_DWML_MAIN, NS_DWML_PIC, NS_REL, NS_WPD_ML, NS_WP_ML,
};
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
    let payload = document_xml.as_bytes();
    replace_file_in_zip(zip_payload, "word/document.xml", payload)
}

pub(crate) fn replace_file_in_zip(
    zip_payload: &mut ZipArchive<Cursor<Vec<u8>>>,
    file_path_in_zip: &str,
    file_payload: &[u8],
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
        let excluded_path = Path::new(file_path_in_zip);

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
        zip.write_all(&file_payload)?;
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

pub(crate) fn owned_name(
    prefix: &Option<String>,
    ns: &Option<String>,
    tag_name: &String,
) -> xml::name::OwnedName {
    xml::name::OwnedName {
        local_name: tag_name.clone(),
        namespace: ns.clone(),
        prefix: prefix.clone(),
    }
}

pub(crate) fn owned_attribute(
    prefix: &Option<String>,
    ns: &Option<String>,
    name: &String,
    value: &String,
) -> xml::attribute::OwnedAttribute {
    xml::attribute::OwnedAttribute {
        name: owned_name(prefix, ns, name),
        value: value.clone(),
    }
}

pub(crate) fn start_tag_event(
    prefix: &Option<String>,
    namespace: &Option<String>,
    tag_name: &String,
    attrs: Option<&[xml::attribute::OwnedAttribute]>,
) -> xml::reader::XmlEvent {
    let mut ns: BTreeMap<String, String> = BTreeMap::new();

    if let Some(ns_str) = namespace {
        if let Some(pref_str) = prefix {
            ns.insert(pref_str.clone(), ns_str.clone());
        }
    }

    let attr_vec = if let Some(attr_array) = attrs {
        Vec::from(attr_array)
    } else {
        vec![]
    };

    xml::reader::XmlEvent::StartElement {
        name: owned_name(prefix, namespace, tag_name),
        namespace: xml::namespace::Namespace(ns),
        attributes: attr_vec,
    }
}

pub(crate) fn end_tag_event(
    prefix: &Option<String>,
    namespace: &Option<String>,
    tag_name: &String,
) -> xml::reader::XmlEvent {
    xml::reader::XmlEvent::EndElement {
        name: owned_name(prefix, namespace, tag_name),
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
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("pPr"),
                None,
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rPr"),
                None,
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event(
            //     "rFonts",
            //     Some(&[("w", "ascii", "Consolas"), ("w", "hAnsi", "Consolas")]),
            // ),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rFonts"),
                Some(&vec![
                    owned_attribute(
                        &Some(String::from("w")),
                        &Some(String::from(NS_WP_ML)),
                        &String::from("ascii"),
                        &String::from("Consolas"),
                    ),
                    owned_attribute(
                        &Some(String::from("w")),
                        &Some(String::from(NS_WP_ML)),
                        &String::from("hAnsi"),
                        &String::from("Consolas"),
                    ),
                ]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rFonts"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event("sz", Some(&[("w", "val", "16")])),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("sz"),
                Some(&vec![owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("val"),
                    &String::from("16"),
                )]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("sz"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("sz"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event("szCs", Some(&[("w", "val", "16")])),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("szCs"),
                Some(&vec![owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("val"),
                    &String::from("16"),
                )]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("szCs"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("szCs"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("rPr"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rPr"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("pPr"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("pPr"),
            ),
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
            // xml_reader_event: start_tag_event("rPr", None),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rPr"),
                None,
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event(
            //     "rFonts",
            //     Some(&[("w", "ascii", "Consolas"), ("w", "hAnsi", "Consolas")]),
            // ),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rFonts"),
                Some(&vec![
                    owned_attribute(
                        &Some(String::from("w")),
                        &Some(String::from(NS_WP_ML)),
                        &String::from("ascii"),
                        &String::from("Consolas"),
                    ),
                    owned_attribute(
                        &Some(String::from("w")),
                        &Some(String::from(NS_WP_ML)),
                        &String::from("hAnsi"),
                        &String::from("Consolas"),
                    ),
                ]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("rFonts"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rFonts"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event("sz", Some(&[("w", "val", "16")])),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("sz"),
                Some(&vec![owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("val"),
                    &String::from("16"),
                )]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("sz"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("sz"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event("szCs", Some(&[("w", "val", "16")])),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("szCs"),
                Some(&vec![owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("val"),
                    &String::from("16"),
                )]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("szCs"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("szCs"),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("rPr"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("rPr"),
            ),
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
        // xml_reader_event: start_tag_event("p", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("p"),
            None,
        ),
        token_text: None,
    });

    result
}

pub(crate) fn heading_prequel_tokens(heading_style: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("p", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("p"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("pPr"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("pStyle"),
            // Some(&[("w", "val", heading_style)]),
            Some(&vec![owned_attribute(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("val"),
                &String::from(heading_style),
            )]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("pStyle"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("pStyle"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("pPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("pPr"),
        ),
        token_text: None,
    });

    result
}

pub(crate) fn paragraph_sequel_tokens() -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    // </w:p>
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("p"),
        ),
        token_text: None,
    });

    result
}

pub(crate) fn heading_sequel_tokens() -> Vec<Token> {
    paragraph_sequel_tokens()
}

pub(crate) fn char_text_tokens(contents: &str, preserve_space: bool) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let attrs_final: Vec<xml::attribute::OwnedAttribute>;
    let attrs: Option<&[xml::attribute::OwnedAttribute]> = if preserve_space {
        // Some(&[("xml", "space", "preserve")])
        attrs_final = vec![owned_attribute(
            &Some(String::from("xml")),
            &None,
            &String::from("space"),
            &String::from("preserve"),
        )];
        Some(&attrs_final)
    } else {
        None
    };

    // <w:t>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("t", attrs),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("t"),
            attrs,
        ),
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
        // xml_reader_event: end_tag_event("t"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("t"),
        ),
        token_text: None,
    });

    result
}

pub(crate) fn run_start_token() -> Token {
    Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("r", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("r"),
            None,
        ),
        token_text: None,
    }
}

pub(crate) fn run_end_token() -> Token {
    Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("r"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("r"),
        ),
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

pub(crate) fn image_paragraph_tokens(
    relationship_id: &str,
    width: u32,
    height: u32,
    serial_number_in_document: usize,
) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::with_capacity(60);

    let width_emu_attr = format!("{}", pixels_to_word_emu(width));
    let height_emu_attr = format!("{}", pixels_to_word_emu(height));

    let serial_id = 10_000 + serial_number_in_document; // Offset the IDs by 10'000 - a dumb way to "ensure" no conflicts with existing
                                                        // serial ids.
    let figure_name = format!("Figure {}", serial_number_in_document);

    result.extend(paragraph_prequel_tokens());
    result.push(run_start_token());
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("rPr"),
            None,
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("noProof", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("noProof"),
            None,
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("noProof"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("noProof"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("rPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("rPr"),
        ),
        token_text: None,
    });

    // <w:drawing>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("drawing", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("drawing"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("inline"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("distT"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("distB"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("distL"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("distR"), &String::from("0")),
            ]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("extent"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("cx"), &width_emu_attr),
                owned_attribute(&None, &None, &String::from("cy"), &height_emu_attr),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("extent"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("effectExtent"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("l"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("t"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("r"), &String::from("3810")),
                owned_attribute(&None, &None, &String::from("b"), &String::from("3810")),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("wp", "effectExtent"),
        xml_reader_event: end_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("effectExtent"),
        ),
        token_text: None,
    });

    // TODO THIS NEEDS TO BE UPDATED
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("docPr"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("id"), &format!("{}", serial_id)),
                owned_attribute(&None, &None, &String::from("name"), &figure_name),
            ]),
        ), // TODO ! Id is weird here
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("wp", "docPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("docPr"),
        ),
        token_text: None,
    });

    // result.push(Token {
    //     token_type: TokenType::Normal,
    //     xml_reader_event: start_tag_event_with_ns("wp", NS_WPD_ML, "cNvGraphicFramePr", None),
    //     token_text: None,
    // });

    // result.push(Token {
    //     token_type: TokenType::Normal,
    //     xml_reader_event: start_tag_event_with_ns(
    //         "a",
    //         NS_DWML_MAIN,
    //         "graphicFrameLocks",
    //         Some(&[("a", "noChangeAspect", "1")]),
    //     ),
    //     token_text: None,
    // });

    // result.push(Token {
    //     token_type: TokenType::Normal,
    //     xml_reader_event: end_tag_event_with_ns("a", "graphicFrameLocks"),
    //     token_text: None,
    // });

    // result.push(Token {
    //     token_type: TokenType::Normal,
    //     xml_reader_event: end_tag_event_with_ns("wp", "cNvGraphicFramePr"),
    //     token_text: None,
    // });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("a", NS_DWML_MAIN, "graphic", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("graphic"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "graphicData",
        //     Some(&[("a", "uri", NS_DWML_PIC)]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("graphicData"),
            Some(&vec![owned_attribute(
                &None,
                &None,
                &String::from("uri"),
                &String::from(NS_DWML_PIC),
            )]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("pic", NS_DWML_PIC, "pic", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("pic"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("pic", NS_DWML_PIC, "nvPicPr", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("nvPicPr"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "pic",
        //     NS_DWML_PIC,
        //     "cNvPr",
        //     Some(&[("pic", "id", "0"), ("pic", "name", "Picture 1")]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("cNvPr"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("id"), &String::from("1")),
                owned_attribute(
                    &None,
                    &None,
                    &String::from("name"),
                    &String::from("Picture 1"),
                ),
            ]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "cNvPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("cNvPr"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("pic", NS_DWML_PIC, "cNvPicPr", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("cNvPicPr"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "picLocks",
        //     Some(&[
        //         ("a", "noChangeAspect", "1"),
        //         ("a", "nowChangeArrowheads", "1"),
        //     ]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("picLocks"),
            Some(&vec![
                owned_attribute(
                    &None,
                    &None,
                    &String::from("noChangeAspect"),
                    &String::from("1"),
                ),
                owned_attribute(
                    &None,
                    &None,
                    &String::from("noChangeArrowheads"),
                    &String::from("1"),
                ),
            ]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "picLocks"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("picLocks"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "cNvPicPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("cNvPicPr"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "nvPicPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("nvPicPr"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("pic", NS_DWML_PIC, "blipFill", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("blipFill"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "blip",
        //     Some(&[("r", "embed", relationship_id)]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("blip"),
            Some(&vec![owned_attribute(
                &Some(String::from("r")),
                &Some(String::from(NS_REL)),
                &String::from("embed"),
                &String::from(relationship_id),
            )]),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "blip"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("blip"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("srcRect"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("srcRect"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("stretch"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("fillRect"),
            None,
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("fillRect"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("stretch"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "blipFill"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("blipFill"),
        ),
        token_text: None,
    });

    // <pic:spPr bwMode="auto"> super important!
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "pic",
        //     NS_DWML_PIC,
        //     "spPr",
        //     Some(&[("pic", "bwMode", "auto")]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("spPr"),
            Some(&vec![owned_attribute(
                &None,
                &None,
                &String::from("bwMode"),
                &String::from("auto"),
            )]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("a", NS_DWML_MAIN, "xfrm", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("xfrm"),
            None,
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "off",
        //     Some(&[("a", "x", "0"), ("a", "y", "0")]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("off"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("x"), &String::from("0")),
                owned_attribute(&None, &None, &String::from("y"), &String::from("0")),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "off"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("off"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "ext",
        //     Some(&[("a", "cx", &width_emu_attr), ("a", "cy", &height_emu_attr)]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("ext"),
            Some(&vec![
                owned_attribute(&None, &None, &String::from("cx"), &width_emu_attr),
                owned_attribute(&None, &None, &String::from("cy"), &height_emu_attr),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "ext"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("ext"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "xfrm"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("xfrm"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns(
        //     "a",
        //     NS_DWML_MAIN,
        //     "prstGeom",
        //     Some(&[("a", "prst", "rect")]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("prstGeom"),
            Some(&vec![owned_attribute(
                &None,
                &None,
                &String::from("prst"),
                &String::from("rect"),
            )]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "prstGeom"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("prstGeom"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event_with_ns("a", NS_DWML_MAIN, "noFill", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("noFill"),
            None,
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "noFill"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("noFill"),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "spPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("spPr"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("pic", "pic"),
        xml_reader_event: end_tag_event(
            &Some(String::from("pic")),
            &Some(String::from(NS_DWML_PIC)),
            &String::from("pic"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "graphicData"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("graphicData"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("a", "graphic"),
        xml_reader_event: end_tag_event(
            &Some(String::from("a")),
            &Some(String::from(NS_DWML_MAIN)),
            &String::from("graphic"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("wp", "inline"),
        xml_reader_event: end_tag_event(
            &Some(String::from("wp")),
            &Some(String::from(NS_WPD_ML)),
            &String::from("inline"),
        ),
        token_text: None,
    });

    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event_with_ns("w", "drawing"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("drawing"),
        ),
        token_text: None,
    });

    result.push(run_end_token());
    result.extend(paragraph_sequel_tokens());

    //
    result
}

pub(crate) fn get_last_id_number_for_document_xml_rels(
    document_xml_rels_tokens: &[Token],
) -> usize {
    // Figure out the latest relationship number in the ID
    let mut latest_rel_no: usize = 0;
    for prequel_token in document_xml_rels_tokens.iter() {
        match &prequel_token.xml_reader_event {
            xml::reader::XmlEvent::StartElement { attributes, .. } => {
                for attr in attributes.iter() {
                    if attr.name.local_name == "Id" {
                        let number_part_of_the_id = &attr.value[3..];
                        let n = number_part_of_the_id.parse::<usize>().unwrap();
                        if n > latest_rel_no {
                            latest_rel_no = n;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    latest_rel_no
}

pub(crate) fn insert_images_in_document_xml_rels(
    document_xml_rels_tokens: &[Token],
    images: &BTreeMap<String, ImageFileContents>,
) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    let len = document_xml_rels_tokens.len();
    let prequel = &document_xml_rels_tokens[..(len - 2)];
    let sequel = &document_xml_rels_tokens[(len - 2)..];

    // Figure out the latest relationship number in the ID
    let mut latest_rel_no: usize = 0;
    for prequel_token in prequel.iter() {
        match &prequel_token.xml_reader_event {
            xml::reader::XmlEvent::StartElement { attributes, .. } => {
                for attr in attributes.iter() {
                    if attr.name.local_name == "Id" {
                        let number_part_of_the_id = &attr.value[3..];
                        let n = number_part_of_the_id.parse::<usize>().unwrap();
                        if n > latest_rel_no {
                            latest_rel_no = n;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    result.extend(Vec::from(prequel));

    for (rel_id, image_contents) in images.iter() {
        let filename = format!("media/{}", &image_contents.file_contents.filename);
        result.push(Token {
            token_type: TokenType::Normal,
            xml_reader_event: start_tag_event(
                &None, &None,
                &String::from("Relationship"),
                Some(&vec![
                    owned_attribute(&None, &None, &String::from("Id"), &String::from(rel_id)),
                    owned_attribute(&None, &None, &String::from("Type"), &String::from("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")),
                    owned_attribute(&None, &None, &String::from("Target"), &String::from(filename)),
                ])),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event_with_ns("r", "Relationship"),
            xml_reader_event: end_tag_event(&None, &None, &String::from("Relationship")),
            token_text: None,
        })
    }

    result.extend(Vec::from(sequel));
    result
}

pub(crate) fn insert_png_content_type(content_type_tokens: &[Token]) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();
    let prequel = &content_type_tokens[..2];
    let sequel = &content_type_tokens[2..];

    result.extend(Vec::from(prequel));
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: start_tag_event(
            &None,
            &None,
            &String::from("Default"),
            Some(&vec![
                owned_attribute(
                    &None,
                    &None,
                    &String::from("Extension"),
                    &String::from("png"),
                ),
                owned_attribute(
                    &None,
                    &None,
                    &String::from("ContentType"),
                    &String::from("image/png"),
                ),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        xml_reader_event: end_tag_event(&None, &None, &String::from("Default")),
        token_text: None,
    });
    result.extend(Vec::from(sequel));

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
        .perform_indent(false)
        .pad_self_closing(false)
        .create_writer(cursor);

    for item in tokens {
        if let Some(writer_event) = item.xml_reader_event.as_writer_event() {
            // the .write method returns a result, the error value of which is
            // of type xml::writer::emitter::EmitterError, which is private...
            // So here we are just passing along a token TextkitDocxError
            // instead.
            let write_result = writer.write(writer_event);
            if let Err(_) = write_result {
                return Err(TextkitDocxError::FailedWriteXml);
            }
        }
    }

    let result = String::from(std::str::from_utf8(&buf).unwrap());
    Ok(result)
}

fn pixels_to_word_emu(pixels: u32) -> u32 {
    pixels * 914_400 / 72 + 2540
}
