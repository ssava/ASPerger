//! VBScript error types and error-handling primitives.

#[derive(Debug, Clone)]
pub struct VBSError {
    pub code: u16,
    pub message: String,
    pub error_type: VBSErrorType,
}

impl VBSError {
    pub fn new(code: u16, message: String, error_type: VBSErrorType) -> Self {
        VBSError {
            code,
            message,
            error_type,
        }
    }

    pub fn is_exit_for(&self) -> bool {
        matches!(self.error_type, VBSErrorType::ExitFor)
    }

    pub fn is_exit_do(&self) -> bool {
        matches!(self.error_type, VBSErrorType::ExitDo)
    }

    pub fn is_exit_function(&self) -> bool {
        matches!(self.error_type, VBSErrorType::ExitFunction)
    }

    pub fn is_exit_sub(&self) -> bool {
        matches!(self.error_type, VBSErrorType::ExitSub)
    }

    pub fn with_code(mut self, code: i32) -> Self {
        self.code = code as u16;
        self
    }
}

impl std::fmt::Display for VBSError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[Code {}]: {}", self.code, self.message)
    }
}

impl std::error::Error for VBSError {}

/// Categorised VBScript error types used for control flow and error reporting.
/// SyntaxError, ValueError, RuntimeError, NotImplementedError are real errors;
/// ExitFor/ExitDo/ExitFunction/ExitSub are control-flow signals.
#[derive(Debug, Clone, Copy)]
pub enum VBSErrorType {
    SyntaxError = 1001,
    ValueError = 1002,
    RuntimeError = 1003,
    NotImplementedError = 1004,
    ExitFor,
    ExitDo,
    ExitFunction,
    ExitSub,
}

impl VBSErrorType {
    pub fn into_error(self, message: String) -> VBSError {
        VBSError::new(self as u16, message, self)
    }
}
