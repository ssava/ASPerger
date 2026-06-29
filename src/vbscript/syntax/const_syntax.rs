use super::VBSyntax;
use crate::vbscript::compiler::Compiler;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::instruction::Instruction;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

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

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        for (var_name, expr) in &self.var_names {
            compiler.compile_expr(expr);
            let name_lower = var_name.to_lowercase();
            if let Some(slot) = compiler.local_slot(&name_lower) {
                compiler.emit(Instruction::StoreLocal(slot));
            } else {
                let idx = compiler.add_constant(VBValue::String(name_lower.into()));
                compiler.emit(Instruction::StoreGlobal(idx));
            }
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
