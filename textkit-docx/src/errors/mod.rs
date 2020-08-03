use thiserror::Error;

#[derive(Error, Debug)]
pub enum TextkitDocxError {
    #[error("Docx file is malformed.")]
    Malformed {
        #[from]
        source: xml::reader::Error,
    },

    #[error(transparent)]
    DocxZip {
        #[from]
        source: zip::result::ZipError,
    },

    #[error(transparent)]
    DocxIo {
        #[from]
        source: std::io::Error,
    },
}
