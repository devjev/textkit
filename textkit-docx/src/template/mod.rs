//! This module is responsible for implementation of template functionality
//! in Docx files.

use crate::{
    errors::TextkitDocxError,
    parse::{find_template_areas, parse_page_dimensions, unzip_text_file, xml_to_token_vec},
    render::{
        datakit_table::datakit_table_to_tokens, get_last_id_number_for_document_xml_rels,
        insert_images_in_document_xml_rels, insert_png_content_type, jupyter_nb::*,
        markdown::markdown_to_tokens, new_zip_bytes_with_document_xml, render_and_paste_tokens,
        replace_file_in_zip, write_token_vector_to_string,
    },
    DocxPayload, ImageFileContents, PageDimensions, TemplateArea, TemplatePlaceholder, Token,
    TokenType, PAT_HB_ALL,
};
use datakit::table::Table;
use regex::Regex;
use serde::Serialize;
use std::collections::{BTreeMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::PathBuf;
use zip::ZipArchive;

/// A .docx template supporting Handlebars syntax.
#[derive(Debug)]
pub struct DocxTemplate {
    source_payload: DocxPayload,
    document_xml: String,
    document_xml_rels: String,
    content_types: String,
    tokens: Vec<Token>,
    document_rels_tokens: Vec<Token>,
    content_types_tokens: Vec<Token>,
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
        let document_xml_rels =
            unzip_text_file(&mut source_payload, "word/_rels/document.xml.rels")?;
        let content_types = unzip_text_file(&mut source_payload, "[Content_Types].xml")?;

        let tokens = xml_to_token_vec(&document_xml)?;
        let document_rels_tokens = xml_to_token_vec(&document_xml_rels)?;
        let content_types_tokens = xml_to_token_vec(&content_types)?;
        let dimensions = parse_page_dimensions(&document_xml)?;
        let template_areas = find_template_areas(&tokens, "p");

        //println!("template_areas = {:#?}", template_areas);

        Ok(Self {
            source_payload,
            document_xml,
            document_xml_rels,
            content_types,
            tokens,
            document_rels_tokens,
            content_types_tokens,
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

        // Here we track all possible images that need to be added to the DOCX file
        // via templating (for example, by importing a Jupyter Notebook with charts).
        // To add images to a DOCX file, not only do we need to modify the `word/document.xml`
        // file, but also the `word/_rels/document.xml.rels`, as well as adding the file to
        // `media/<filename>.png`.
        let mut images: BTreeMap<String, ImageFileContents> = BTreeMap::new();

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

        // Figure out the latest numerical part of the IDs in word/_rels/document.xml.rels
        let mut latest_rels_id =
            get_last_id_number_for_document_xml_rels(&self.document_rels_tokens);

        // Also, we need a json serialized version of the data (mimicking Handlebars)
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
                                    } else if helper_name == "jupyter" {
                                        if let Some(jupyter_nb) =
                                            serialized_data.get(&placeholder.expression)
                                        {
                                            let notebook: JupyterNotebook =
                                                serde_json::from_value(jupyter_nb.clone())?;
                                            let notebook_tokens = jupyter_nb_to_tokens(
                                                &notebook,
                                                &mut latest_rels_id,
                                                &mut images,
                                            );
                                            result.extend(notebook_tokens);
                                        }
                                    } else if helper_name == "markdown" {
                                        if let Some(markdown_source) =
                                            serialized_data.get(&placeholder.expression)
                                        {
                                            let source_text: String =
                                                serde_json::from_value(markdown_source.clone())?;
                                            let md_tokens = markdown_to_tokens(&source_text);
                                            result.extend(md_tokens);
                                        }
                                    } else {
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

        // Deal with any potential images that need to be inserted as well.
        let new_document_xml_rels_tokens =
            insert_images_in_document_xml_rels(&self.document_rels_tokens, &images);

        let document_xml_rels_contents =
            write_token_vector_to_string(&new_document_xml_rels_tokens)?;

        let mut new_payload = replace_file_in_zip(
            &mut payload,
            "word/_rels/document.xml.rels",
            &document_xml_rels_contents.as_bytes(),
        )?;
        let mut cursor = Cursor::new(new_payload);
        let mut new_zip = ZipArchive::new(cursor)?;

        for (_, image_contents) in images.iter() {
            let path_to_image = format!("word/media/{}", image_contents.file_contents.filename);
            new_payload = replace_file_in_zip(
                &mut new_zip,
                &path_to_image,
                &image_contents.file_contents.payload,
            )?;
            cursor = Cursor::new(new_payload);
            new_zip = ZipArchive::new(cursor)?;
        }

        let new_content_type_tokens = insert_png_content_type(&self.content_types_tokens);
        let new_content_type_payload = write_token_vector_to_string(&new_content_type_tokens)?;
        new_payload = replace_file_in_zip(
            &mut new_zip,
            "[Content_Types].xml",
            new_content_type_payload.as_bytes(),
        )?;

        cursor = Cursor::new(new_payload);
        new_zip = ZipArchive::new(cursor)?;

        new_zip_bytes_with_document_xml(&mut new_zip, &document_xml_contents)
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
