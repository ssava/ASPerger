use std::sync::Arc;

use crate::asp::parser::AspBlock;
use crate::vbscript::ExecutionContext;

use super::asp_error::ASPError;

/// Defines the interface for all handlers in the chain.
pub trait Handler: Send + Sync {
    /// Sets the next handler in the chain.
    fn set_next(&mut self, next: Arc<dyn Handler + Send + Sync>);

    /// Handles an ASP block. If the handler cannot process the block,
    /// it passes the block to the next handler in the chain.
    fn handle(&self, block: &AspBlock, context: &mut ExecutionContext) -> Result<(), ASPError>;
}