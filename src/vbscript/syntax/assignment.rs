use super::VBSyntax;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

pub struct Assignment {
    var_name: String,
    expr: Expr,
}

impl Assignment {
    pub fn new(var_name: String, expr: Expr) -> Self {
        Assignment { var_name, expr }
    }
}

impl VBSyntax for Assignment {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let value = evaluate(&self.expr, context)?;
        context.set_variable(&self.var_name, value);
        Ok(())
    }
}
