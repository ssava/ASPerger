use super::super::compiler::Compiler;
use super::super::execution_context::ErrorMode;
use super::super::instruction::Instruction;
use super::super::vbs_error::VBSError;
use super::super::ExecutionContext;
use super::VBSyntax;

/// AST node for `On Error Resume Next`.
///
/// Switches the interpreter into "resume next" mode where runtime errors
/// are silently recorded in `Err.Number` / `Err.Description` instead of
/// halting execution.
#[derive(Clone)]
pub struct OnErrorResumeNext;

impl VBSyntax for OnErrorResumeNext {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        context.set_error_mode(ErrorMode::ResumeNext);
        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        compiler.emit(Instruction::OnErrorResumeNext);
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}

/// AST node for `On Error GoTo 0`.
///
/// Restores normal error handling (errors halt execution)
/// and clears any recorded error state.
#[derive(Clone)]
pub struct OnErrorGoto0;

impl VBSyntax for OnErrorGoto0 {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        context.set_error_mode(ErrorMode::Normal);
        context.clear_err();
        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        compiler.emit(Instruction::OnErrorGoto0);
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
