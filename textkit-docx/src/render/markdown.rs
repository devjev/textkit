use crate::{
    render::{
        char_text_tokens, end_tag_event, heading_prequel_tokens, heading_sequel_tokens,
        paragraph_prequel_tokens, paragraph_sequel_tokens, run_end_token, run_start_token,
        start_tag_event,
    },
    Token, TokenType, NS_WP_ML,
};
use pulldown_cmark::{Options, Parser};

pub(crate) fn markdown_to_tokens(md_text: &str) -> Vec<Token> {
    let mut result: Vec<Token> = Vec::new();

    let mut markdown_options = Options::empty();
    markdown_options.insert(Options::ENABLE_STRIKETHROUGH);

    let parser = Parser::new_ext(md_text, markdown_options);
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
                // xml_reader_event: start_tag_event("rPr", None),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("rPr"),
                    None,
                ),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: start_tag_event("i", None),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("i"),
                    None,
                ),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("i"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("i"),
                ),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: start_tag_event("iCs", None),
                xml_reader_event: start_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("iCs"),
                    None,
                ),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("iCs"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("iCs"),
                ),
                token_text: None,
            });
            emphasis_start.push(Token {
                token_type: TokenType::Normal,
                // xml_reader_event: end_tag_event("rPr"),
                xml_reader_event: end_tag_event(
                    &Some(String::from("w")),
                    &Some(String::from(NS_WP_ML)),
                    &String::from("rPr"),
                ),
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
