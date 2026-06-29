//! Syntax node trait and AST node types for parsed VBScript constructs.
//! Each variant (assignment, method call, dim, redim, etc.) implements
//! the `VBSyntax` trait and is executed during script evaluation.

pub trait VBSyntax {
    fn execute(&self, context: &mut crate::vbscript::ExecutionContext) -> Result<(), VBSError>;
    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError>;
    fn clone_box(&self) -> Box<dyn VBSyntax>;
}

// Re-export all syntax constructs
mod array_assignment;
mod assignment;
mod const_syntax;
mod dim;
mod erase;
mod method_call;
mod on_error;
mod property_set;
mod redim;
mod response_cookies_set;
mod response_write;
use super::vbs_error::VBSError;
use crate::vbscript::compiler::Compiler;

pub use array_assignment::ArrayAssignment;
pub use assignment::Assignment;
pub use const_syntax::Const;
pub use dim::Dim;
pub use erase::Erase;
pub use method_call::MethodCall;
pub use on_error::{OnErrorGoto0, OnErrorResumeNext};
pub use property_set::PropertySet;
pub use redim::ReDim;
pub use response_cookies_set::{ResponseCookiesSet, ResponseCookiesSetProp};
pub use response_write::ResponseWrite;
