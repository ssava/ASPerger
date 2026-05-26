#[derive(Debug)]
pub struct ASPError {
    pub code: u16,
    pub message: String,
}

impl ASPError {
    pub fn new(code: u16, message: impl Into<String>) -> Self {
        ASPError { code, message: message.into() }
    }
}

impl std::fmt::Display for ASPError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[Code {}]: {}", self.code, self.message)
    }
}

impl std::error::Error for ASPError {}