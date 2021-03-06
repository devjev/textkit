use thiserror::Error;

#[derive(Error, Debug)]
pub enum TextkitDocxError {
    #[error("Unable to parse input data for template")]
    BadInputData {
        #[from]
        source: serde_json::error::Error,
    },

    #[error("Malformed document")]
    Malformed(String),

    #[error("Failed to write XML data")]
    FailedWriteXml,

    #[error("Docx file is malformed.")]
    FailedReadXml {
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
