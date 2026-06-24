//! ASP-level error type used throughout the server layer.

/// HTTP-level error used across the server layer.
///
/// Unlike `VBSError` (which carries VBScript-level error codes and
/// control-flow signals), `ASPError` represents an HTTP response status
/// with a human-readable message intended for the client.
#[derive(Debug)]
pub struct ASPError {
    /// HTTP status code (e.g. 404, 403, 500).
    pub code: u16,
    /// Human-readable error description.
    pub message: String,
}

impl ASPError {
    /// Create a new ASP error with the given HTTP status code and message.
    pub fn new(code: u16, message: impl Into<String>) -> Self {
        ASPError {
            code,
            message: message.into(),
        }
    }
}

impl ASPError {
    /// Render the error as a styled HTML page suitable for returning to the browser.
    pub fn render_html(&self) -> String {
        format!(
            "<html><head><title>ASP Error {}</title><style>\
             body{{font-family:monospace;background:#f8f8f8;padding:2em}}\
             h1{{color:#c00;font-size:1.5em}}\
             .code{{color:#666}}\
             .msg{{background:#fff;border:1px solid #ddd;padding:1em;margin:1em 0}}\
             </style></head><body>\
             <h1>ASP Error</h1>\
             <p class=\"code\">HTTP {}</p>\
             <div class=\"msg\">{}</div>\
             </body></html>",
            self.code, self.code, self.message
        )
    }
}

impl std::fmt::Display for ASPError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "[Code {}]: {}", self.code, self.message)
    }
}

impl std::error::Error for ASPError {}
