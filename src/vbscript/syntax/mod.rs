// Define the VBSyntax trait
pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), VBSError>;
}

// Re-export all syntax constructs
mod response_write;
mod dim;
mod assignment;
mod method_call;
mod redim;
mod array_assignment;
mod property_set;

pub use response_write::ResponseWrite;
use super::vbs_error::VBSError;
pub use dim::Dim;
pub use assignment::Assignment;
pub use method_call::MethodCall;
pub use redim::ReDim;
pub use array_assignment::ArrayAssignment;
pub use property_set::PropertySet;