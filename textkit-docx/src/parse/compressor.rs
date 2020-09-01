// This is a temporary solution - should be done properly in the accumulator
// FSM.

use crate::render::owned_name;
use crate::{Token, TokenType, NS_WP_ML};
use xml::reader::XmlEvent;

pub(crate) fn compress_tokens(tokens: &Vec<Token>) {
    let mut result: Vec<Token> = Vec::new();
    let mut inside_paragraph = false;
    let mut compressing_template = false;

    let p_name = owned_name(
        &Some(String::from("w")),
        &Some(String::from(NS_WP_ML)),
        &String::from("p"),
    );

    for token in tokens.iter() {
        match &token.xml_reader_event {
            XmlEvent::StartElement { name, .. } if name == &p_name => {
                inside_paragraph = true;
            }
            XmlEvent::StartElement { name, .. } if name != &p_name && inside_paragraph => {
                //println!("{:#?}", token); // accumulate prequel
            }
            _ => {}
        }

        // iterate until a <w:p> is reached,
        // mark that down, and iterate until Template.
        // markd that down and untill there's </w:p> compress all
        // characters into a single token.
    }
}
