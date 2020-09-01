use crate::render::{end_tag_event, owned_attribute, paragraph_tokens, start_tag_event};
use crate::NS_WP_ML;
use crate::{PageDimensions, Token, TokenType};
use datakit::{table::Table, value::definitions::*, value::primitives::*};

pub(crate) fn datakit_table_to_tokens(table: &Table, dims: &PageDimensions) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    // <w:tbl>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("tbl", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tbl"),
            None,
        ),
        token_text: None,
    });

    // <w:tblPr>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("tblPr", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblPr"),
            None,
        ),
        token_text: None,
    });

    // <w:tblStyle w:val="TableGrid" />
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("tblStyle", Some(&[("w", "val", "TableGrid")])),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblStyle"),
            Some(&vec![owned_attribute(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("val"),
                &String::from("TableGrid"),
            )]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tblStyle"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblStyle"),
        ),
        token_text: None,
    });

    // <w:tblW w:w="0" w:type="auto" />
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("tblW", Some(&[("w", "w", "0"), ("w", "type", "auto")])),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblW"),
            Some(&vec![
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("w"),
                    &String::from("0"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("type"),
                    &String::from("auto"),
                ),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tblW"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblW"),
        ),
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
        // xml_reader_event: start_tag_event(
        //     "tblLook",
        //     Some(&[
        //         ("w", "val", "04A0"),
        //         ("w", "firstRow", "1"),
        //         ("w", "lastRow", "0"),
        //         ("w", "firstColumn", "1"),
        //         ("w", "lastColumn", "0"),
        //         ("w", "noHBand", "0"),
        //         ("w", "noVBand", "1"),
        //     ]),
        // ),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblLook"),
            Some(&vec![
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("val"),
                    &String::from("04A0"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("firstRow"),
                    &String::from("1"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("lastRow"),
                    &String::from("0"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("firstColumn"),
                    &String::from("1"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("lastColumn"),
                    &String::from("0"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("noHBand"),
                    &String::from("0"),
                ),
                owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("noVBand"),
                    &String::from("0"),
                ),
            ]),
        ),
        token_text: None,
    });
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tblLook"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblLook"),
        ),
        token_text: None,
    });

    // </w:tblPr>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tblPr"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblPr"),
        ),
        token_text: None,
    });

    // <w:tblGrid>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: start_tag_event("tblGrid", None),
        xml_reader_event: start_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblGrid"),
            None,
        ),
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
            // xml_reader_event: start_tag_event("gridCol", Some(&[("w", "w", &col_width_str)])),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("gridCol"),
                Some(&vec![owned_attribute(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("w"),
                    &col_width_str,
                )]),
            ),
            token_text: None,
        });
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("gridCol"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("gridCol"),
            ),
            token_text: None,
        });
    }

    // <w:tblGrid>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tblGrid"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tblGrid"),
        ),
        token_text: None,
    });

    // Now we need to populate the contents of the the table
    for row_i in 0..table.len() {
        // <w:tr> - we deliberately omitting any kind of id attributes (like w:rsidR).
        // That's not very compliant as far as I know, but MS Word handles it pretty
        // well, so IDs are a TODO for future versions.
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: start_tag_event("tr", None),
            xml_reader_event: start_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("tr"),
                None,
            ),
            token_text: None,
        });

        // Populate all table cells for the current row.
        for col_i in 0..no_of_cols {
            // <w:tc>
            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: start_tag_event("tc", None),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tc"),
                    None,
                ),
                token_text: None,
            });

            // <w:tcPr>
            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: start_tag_event("tcPr", None),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tcPr"),
                    None,
                ),
                token_text: None,
            });

            // <w:tcW w:w="<COL_WIDTH>" w:type="dxa" />
            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: start_tag_event(
                //     "tcW",
                //     Some(&[("w", "w", &col_width_str), ("w", "type", "dxa")]),
                // ),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tcW"),
                    Some(&vec![
                        owned_attribute(
                            &Some(String::from("w")),
                            &Some(String::from(NS_WP_ML)),
                            &String::from("w"),
                            &col_width_str,
                        ),
                        owned_attribute(
                            &Some(String::from("w")),
                            &Some(String::from(NS_WP_ML)),
                            &String::from("type"),
                            &String::from("dxa"),
                        ),
                    ]),
                ),
                token_text: None,
            });
            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("tcW"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tcW"),
                ),
                token_text: None,
            });

            // </w:tcPr>
            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("tcPr"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tcPr"),
                ),
                token_text: None,
            });

            // TODO fill this thing with data from the value
            let cell_value = &table.columns()[col_i][row_i];

            match cell_value {
                Value::Text(text) => {
                    let paragraph_tokens = paragraph_tokens(text);
                    result.extend(paragraph_tokens);
                }
                Value::Number(Numeric::Integer(int)) => {
                    let int_str = format!("{}", int);
                    let paragraph_tokens = paragraph_tokens(&int_str);
                    result.extend(paragraph_tokens);
                }
                Value::Number(Numeric::Real(real)) => {
                    let real_str = format!("{:.3}", real);
                    let paragraph_tokens = paragraph_tokens(&real_str);
                    result.extend(paragraph_tokens);
                }
                Value::Boolean(boolean) => {
                    let bool_str = if *boolean { "Yes" } else { "No" };
                    let paragraph_tokens = paragraph_tokens(&bool_str);
                    result.extend(paragraph_tokens);
                }
                Value::DateTime(dt) => {
                    let dt_str = format!("{}", dt);
                    let paragraph_tokens = paragraph_tokens(&dt_str);
                    result.extend(paragraph_tokens);
                }
                // TODO implement the rest of value types
                _ => (),
            }

            result.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("tc"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("tc"),
                ),
                token_text: None,
            });
        }

        // </w:tr>
        result.push(Token {
            token_type: TokenType::Normal,
            // xml_reader_event: end_tag_event("tr"),
            xml_reader_event: end_tag_event(
                &Some(String::from("w")),
                &Some(String::from(NS_WP_ML)),
                &String::from("tr"),
            ),
            token_text: None,
        });
    }

    // </w:tbl>
    result.push(Token {
        token_type: TokenType::Normal,
        // xml_reader_event: end_tag_event("tbl"),
        xml_reader_event: end_tag_event(
            &Some(String::from("w")),
            &Some(String::from(NS_WP_ML)),
            &String::from("tbl"),
        ),
        token_text: None,
    });

    result
}
