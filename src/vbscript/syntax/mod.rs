// Define the VBSyntax trait
pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), VBSError>;
}

// Re-export all syntax constructs
mod response_write;
mod dim;
mod assignment;
mod if_statement;
mod for_loop;
mod while_loop;
mod function_decl;
mod call_function;

pub use response_write::ResponseWrite;
use super::vbs_error::VBSError;
pub use dim::Dim;
pub use assignment::Assignment;