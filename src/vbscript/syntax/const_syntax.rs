use super::VBSyntax;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

/// AST node for `Const name = expr` declarations.
///
/// Evaluates the expression at parse time (during the first execution
/// pass) and sets the constant as a variable in the current scope.
/// Note: VBScript `Const` values are immutable, but the interpreter
/// does not enforce immutability here.
#[derive(Clone)]
pub struct Const {
    var_names: Vec<(String, Expr)>,
}

impl Const {
    pub fn new(var_names: Vec<(String, Expr)>) -> Self {
        Const { var_names }
    }
}

impl VBSyntax for Const {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        for (var_name, expr) in &self.var_names {
            let value = evaluate(expr, context)?;
            context.set_variable(var_name, value);
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
