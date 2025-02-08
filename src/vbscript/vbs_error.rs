#[derive(Debug)]
pub struct VBSError {
    pub code: u32,
    pub message: String,
}

impl VBSError {
    pub fn new(code: u32, message: String) -> Self {
        VBSError { code, message }
    }

    pub fn to_string(&self) -> String {
        format!("[Codice {}]: {}", self.code, self.message)
    }
}

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
        VBSError::new(self as u32, message)
    }
}