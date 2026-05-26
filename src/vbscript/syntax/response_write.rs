use super::VBSyntax;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

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
}
