use clap::Clap;
use std::fs::File;
use std::io::{Error, ErrorKind, Write};
use std::path::PathBuf;
use textkit_docx::DocxTemplate;

fn main() -> std::io::Result<()> {
    let opts = Options::parse();
    match opts.subcmd {
        SubCommand::DocxTemplate(subcmd_opts) => docx_template_apply(subcmd_opts),
    }
}

fn docx_template_apply(opts: DocxTemplateOptions) -> std::io::Result<()> {
    if opts.output == opts.template {
        let error = Error::new(
            ErrorKind::InvalidInput,
            "Output file name cannot be the same as the template file name.",
        );
        Err(error)
    } else if opts.output == opts.json {
        let error = Error::new(
            ErrorKind::InvalidInput,
            "Output file name cannot be the same as the JSON file name.",
        );
        Err(error)
    } else {
        if let Ok(template) = DocxTemplate::from_file(&opts.template) {
            let data_fh = File::open(&opts.json)?;
            let mut output_fh = File::create(&opts.output)?;

            let data: serde_json::Value = serde_json::from_reader(data_fh)?;

            if let Ok(mut new_docx_data) = template.render(&data) {
                output_fh.write_all(&mut new_docx_data)?;
            } else {
                let error = Error::new(ErrorKind::Other, "Could not render the docx template!");
                return Err(error);
            }

            Ok(())
        } else {
            let error = Error::new(ErrorKind::InvalidData, "Could not read the DOCX template.");
            Err(error)
        }
    }
}

#[derive(Clap, Debug)]
#[clap(version = "0.1.0", author = "Jevgeni Tarasov <jevgeni@hey.com>")]
struct Options {
    #[clap(subcommand)]
    subcmd: SubCommand,
}

#[derive(Clap, Debug)]
enum SubCommand {
    DocxTemplate(DocxTemplateOptions),
}

/// Apply JSON data to a Textkit DOCX template.
#[derive(Clap, Debug)]
struct DocxTemplateOptions {
    /// Path to the .docx file acting as a template.
    #[clap(short, long)]
    template: PathBuf,

    /// Path to the .json file with the data to be pasted into the template.
    #[clap(short, long)]
    json: PathBuf,

    /// Output file name.
    #[clap(short, long)]
    output: PathBuf,
}
