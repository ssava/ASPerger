use std::sync::Arc;

use crate::vbscript::ExecutionContext;

use super::{asp_error::ASPError, handler::Handler, parser::AspBlock};

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