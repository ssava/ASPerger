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
pub use dim::Dim;
pub use assignment::Assignment;
pub use if_statement::IfStatement;
pub use for_loop::ForLoop;
pub use while_loop::WhileLoop;
pub use function_decl::Function;
pub use call_function::CallFunction;

use super::vbs_error::VBSError;