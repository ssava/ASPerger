// Define the VBSyntax trait
pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), String>;
}

// Re-export all syntax constructs
mod response_write;
mod dim;
mod assignment;
mod if_statement;

pub use response_write::ResponseWrite;
pub use dim::Dim;
pub use assignment::Assignment;
pub use if_statement::IfStatement;