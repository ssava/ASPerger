#[derive(Debug)]
pub struct VBSError {
    pub code: u16,
    pub message: String,
}

impl VBSError {
    pub fn new(code: u16, message: String) -> Self {
        VBSError { code, message }
    }
}

impl std::fmt::Display for VBSError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[Codice {}]: {}", self.code, self.message)
    }
}

impl std::error::Error for VBSError {}

#[derive(Debug)]
pub enum VBSErrorType {
    SyntaxError = 1001,
    TypeError = 1002,
    NameError = 1003,
    ValueError = 1004,
    RuntimeError = 1005,
    NotImplementedError = 1006,
    BlockMismatchError = 1007,
}

impl VBSErrorType {
    pub fn into_error(self, message: String) -> VBSError {
        VBSError::new(self as u16, message)
    }
}