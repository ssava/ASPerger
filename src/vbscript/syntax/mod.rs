//! Syntax node trait and AST node types for parsed VBScript constructs.
//! Each variant (assignment, method call, dim, redim, etc.) implements
//! the `VBSyntax` trait and is executed during script evaluation.

pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), VBSError>;
    fn clone_box(&self) -> Box<dyn VBSyntax>;
}

// Re-export all syntax constructs
mod array_assignment;
mod assignment;
mod dim;
mod method_call;
mod on_error;
mod property_set;
mod redim;
mod response_cookies_set;
mod response_write;

use super::vbs_error::VBSError;
pub use array_assignment::ArrayAssignment;
pub use assignment::Assignment;
pub use dim::Dim;
pub use method_call::MethodCall;
pub use on_error::{OnErrorGoto0, OnErrorResumeNext};
pub use property_set::PropertySet;
pub use redim::ReDim;
pub use response_cookies_set::ResponseCookiesSet;
pub use response_write::ResponseWrite;
