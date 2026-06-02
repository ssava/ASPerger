// Define the VBSyntax trait
pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), VBSError>;
}

// Re-export all syntax constructs
mod response_write;
mod response_cookies_set;
mod dim;
mod assignment;
mod method_call;
mod redim;
mod array_assignment;
mod property_set;
mod on_error;

pub use response_write::ResponseWrite;
pub use response_cookies_set::ResponseCookiesSet;
use super::vbs_error::VBSError;
pub use dim::Dim;
pub use assignment::Assignment;
pub use method_call::MethodCall;
pub use redim::ReDim;
pub use array_assignment::ArrayAssignment;
pub use property_set::PropertySet;
pub use on_error::{OnErrorResumeNext, OnErrorGoto0};