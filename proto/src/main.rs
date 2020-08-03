use std::fs::File;
use std::io::{Write, Read};
use std::collections::HashMap;
use textkit_word::DocxTemplate;
use std::path::PathBuf;
use structopt::StructOpt;

#[derive(StructOpt, Debug)]
struct Options {

    /// The path to the Word DOCX file to be used as template.
    #[structopt(short="t", long)]
    pub template: PathBuf,

    /// The text file with the contents that will be placed in the
    /// {{test}} placeholder in the Word template file.
    #[structopt(short="c", long)]
    pub contents: PathBuf,

    /// The file name of the output DOCX file
    #[structopt(short="o", long)]
    pub output: PathBuf,
}

fn main() {
    let options = Options::from_args();

    let docx_template = DocxTemplate::from_file(&options.template).unwrap();
    let mut contents_fh = File::open(options.contents).unwrap();
    let mut contents = String::new();
    contents_fh.read_to_string(&mut contents).unwrap();

    let mut data_context: HashMap<String, String> = HashMap::new();
    data_context.insert("test".into(), contents);

    let new_payload = docx_template.render(&data_context);
    let mut target_fh = File::create(options.output).unwrap();
    target_fh.write(&new_payload).unwrap();
}
