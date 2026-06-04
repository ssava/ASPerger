use super::super::execution_context::ErrorMode;
use super::super::vbs_error::VBSError;
use super::super::ExecutionContext;
use super::VBSyntax;

pub struct OnErrorResumeNext;

impl VBSyntax for OnErrorResumeNext {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        context.set_error_mode(ErrorMode::ResumeNext);
        Ok(())
    }
}

pub struct OnErrorGoto0;

impl VBSyntax for OnErrorGoto0 {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        context.set_error_mode(ErrorMode::Normal);
        context.clear_err();
        Ok(())
    }
}
