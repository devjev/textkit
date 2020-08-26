use crate::render::{make_end_tag_event, make_paragraph_tokens, make_start_tag_event};
use crate::{PageDimensions, Token, TokenType};
use datakit::{table::Table, value::definitions::*, value::primitives::*};

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
        xml_reader_event: make_start_tag_event("tblStyle", Some(&[("w", "val", "TableGrid")])),
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
        xml_reader_event: make_start_tag_event(
            "tblW",
            Some(&[("w", "w", "0"), ("w", "type", "auto")]),
        ),
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
                ("w", "val", "04A0"),
                ("w", "firstRow", "1"),
                ("w", "lastRow", "0"),
                ("w", "firstColumn", "1"),
                ("w", "lastColumn", "0"),
                ("w", "noHBand", "0"),
                ("w", "noVBand", "1"),
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
            xml_reader_event: make_start_tag_event("gridCol", Some(&[("w", "w", &col_width_str)])),
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
                    Some(&[("w", "w", &col_width_str), ("w", "type", "dxa")]),
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
