use std::sync::Arc;

use crate::asp::parser::AspBlock;
use crate::vbscript::{ExecutionContext, Interpreter};

use super::asp_error::ASPError;

/// Defines the interface for all handlers in the chain.
pub trait Handler: Send + Sync {
    /// Sets the next handler in the chain.
    fn set_next(&mut self, next: Arc<dyn Handler + Send + Sync>);

    /// Handles an ASP block. If the handler cannot process the block,
    /// it passes the block to the next handler in the chain.
    fn handle(&self, block: &AspBlock, context: &mut ExecutionContext) -> Result<(), ASPError>;
}

/// Handles `AspBlock::Html` blocks by writing the HTML content to the response buffer.
pub struct HtmlHandler {
    next: Option<Arc<dyn Handler + Send + Sync>>,
}

impl HtmlHandler {
    /// Creates a new `HtmlHandler`.
    pub fn new() -> Self {
        HtmlHandler { next: None }
    }
}

impl Handler for HtmlHandler {
    fn set_next(&mut self, next: Arc<dyn Handler + Send + Sync>) {
        self.next = Some(next);
    }

    fn handle(&self, block: &AspBlock, context: &mut ExecutionContext) -> Result<(), ASPError> {
        match block {
            AspBlock::Html(html) => {
                context.write(html);
                return Ok(());
            }
            _ => {}
        }
        if let Some(next) = &self.next {
            // Pass the block to the next handler in the chain.
            next.handle(block, context)
        } else {
            // No handler available to process the block.
            Err(ASPError::new(
                500,
                "No handler available for the block".to_string(),
            ))
        }
    }
}
/// Handles `AspBlock::Code` blocks by executing VBScript through the interpreter.
pub struct CodeHandler {
    interpreter: Arc<dyn Interpreter>,
    next: Option<Arc<dyn Handler + Send + Sync>>,
}

impl CodeHandler {
    /// Creates a new `CodeHandler` with the given `Interpreter`.
    pub fn new(interpreter: Arc<dyn Interpreter>) -> Self {
        CodeHandler {
            interpreter,
            next: None,
        }
    }
}

impl Handler for CodeHandler {
    fn set_next(&mut self, next: Arc<dyn Handler + Send + Sync>) {
        self.next = Some(next);
    }

    fn handle(&self, block: &AspBlock, context: &mut ExecutionContext) -> Result<(), ASPError> {
        if let AspBlock::Code(code, start_line) = block {
            context.code_start_line = *start_line;
            self.interpreter
                .execute(code, context)
                .map_err(|e| ASPError::new(500, e.to_string()))
        } else if let Some(next) = &self.next {
            // Pass the block to the next handler in the chain.
            next.handle(block, context)
        } else {
            // No handler available to process the block.
            Err(ASPError::new(
                500,
                "No handler available for the block".to_string(),
            ))
        }
    }
}
