use clap::Clap;
use std::fs::File;
use std::io::Read;
use std::io::{Error, ErrorKind, Write};
use std::path::PathBuf;
use textkit_docx::DocxTemplate;

// Jupyter DOMOs
use textkit_docx::render::jupyter_nb::*;

fn main() -> std::io::Result<()> {
    let opts = Options::parse();
    match opts.subcmd {
        SubCommand::DocxTemplate(subcmd_opts) => docx_template_apply(subcmd_opts),
        SubCommand::JupyterConvert(subcmd_opts) => jupyter_to_docx(subcmd_opts),
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
        match DocxTemplate::from_file(&opts.template) {
            Ok(template) => {
                let data_fh = File::open(&opts.json)?;
                let mut output_fh = File::create(&opts.output)?;
                let data: serde_json::Value = serde_json::from_reader(data_fh)?;
                match template.render(&data) {
                    Ok(mut new_docx_data) => {
                        output_fh.write_all(&mut new_docx_data)?;
                    }
                    Err(error) => {
                        let message = format!("Could not render the DOCX template: {:?}", error);
                        let new_error = Error::new(ErrorKind::Other, message);
                        return Err(new_error);
                    }
                }
                Ok(())
            }
            Err(error) => {
                let message = format!("Could not read the DOCX template: {:?}", error);
                let new_error = Error::new(ErrorKind::InvalidData, message);
                Err(new_error)
            }
        }
    }
}

fn jupyter_to_docx(opts: JupyterParsingOptions) -> std::io::Result<()> {
    let mut fh = File::open(opts.jupyter_notebook)?;
    let mut json_payload = String::new();
    fh.read_to_string(&mut json_payload)?;
    let nb: JupyterNotebook = serde_json::from_str(&json_payload)?;
    jupyter_nb_to_tokens_proto(&nb);
    Ok(())
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
    JupyterConvert(JupyterParsingOptions),
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

#[derive(Clap, Debug)]
struct JupyterParsingOptions {
    #[clap(short, long)]
    jupyter_notebook: PathBuf,
}
