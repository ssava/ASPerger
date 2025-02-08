use std::sync::Arc;

use crate::asp::parser::AspBlock;
use crate::vbscript::{ExecutionContext, VBScriptInterpreter};

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
        if let AspBlock::Html(html) = block {
            // Write the HTML content to the response buffer.
            context.write(html);
            Ok(())
        } else if let Some(next) = &self.next {
            // Pass the block to the next handler in the chain.
            next.handle(block, context)
        } else {
            // No handler available to process the block.
            Err(ASPError::new(
                500,
                "Nessun handler disponibile per il blocco".to_string(),
            ))
        }
    }
}
pub struct CodeHandler {
    interpreter: Arc<VBScriptInterpreter>,
    next: Option<Arc<dyn Handler + Send + Sync>>,
}

impl CodeHandler {
    /// Creates a new `CodeHandler` with the given `VBScriptInterpreter`.
    pub fn new(interpreter: Arc<VBScriptInterpreter>) -> Self {
        CodeHandler {
            interpreter,
            next: None,
        }
    }
}

impl Handler for CodeHandler {
    fn set_next(&mut self, next: Arc<dyn Handler + Send + Sync>){
        self.next = Some(next);
    }

    fn handle(&self, block: &AspBlock, context: &mut ExecutionContext) -> Result<(), ASPError> {
        if let AspBlock::Code(code) = block {
            // Execute the VBScript code using the interpreter.
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
                "Nessun handler disponibile per il blocco".to_string(),
            ))
        }
    }
}