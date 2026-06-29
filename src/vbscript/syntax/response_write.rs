use super::VBSyntax;
use crate::vbscript::compiler::Compiler;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::instruction::Instruction;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

/// AST node for `Response.Write expr` statements.
///
/// Evaluates the expression at runtime and appends the string
/// representation to the response buffer.
#[derive(Clone)]
pub struct ResponseWrite {
    expr: Expr,
}

impl ResponseWrite {
    pub fn new(expr: Expr) -> Self {
        ResponseWrite { expr }
    }
}

impl VBSyntax for ResponseWrite {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let value = evaluate(&self.expr, context)?;
        context.write(&value.to_string());
        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        compiler.compile_expr(&self.expr);
        compiler.emit(Instruction::ResponseWrite);
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
