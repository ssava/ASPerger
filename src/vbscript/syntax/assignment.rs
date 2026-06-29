use super::VBSyntax;
use crate::vbscript::compiler::Compiler;
use crate::vbscript::expr::{evaluate, BinOp, Expr};
use crate::vbscript::instruction::Instruction;
use crate::vbscript::value_utils;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

/// AST node for simple variable assignment: `var = expr`.
///
/// Contains an optimisation for the `var = var & expr` pattern (string
/// concatenation to self) that uses `String::push_str` to avoid O(n²)
/// behaviour when building up a string in a loop.
#[derive(Clone)]
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
        // Fast path: detect var = var & expr and use push_str to avoid O(n²)
        if let Expr::BinaryOp { left, op: BinOp::Concat, right } = &self.expr {
            if let Expr::Variable(name) = left.as_ref() {
                if name.eq_ignore_ascii_case(&self.var_name) {
                    let rhs_val = evaluate(right, context)?;
                    let rhs_str = value_utils::to_arg_string(&rhs_val);
                    if let Some(existing) = context.get_variable_mut(&self.var_name) {
                        match existing {
                            VBValue::String(s) => {
                                let mut new_s = s.to_string();
                                new_s.push_str(&rhs_str);
                                *existing = VBValue::String(new_s.into());
                                return Ok(());
                            }
                            VBValue::Empty => {
                                *existing = VBValue::String(rhs_str.into());
                                return Ok(());
                            }
                            _ => {
                                let lhs_str = existing.to_string();
                                *existing = VBValue::String((lhs_str + &rhs_str).into());
                                return Ok(());
                            }
                        }
                    } else {
                        context.set_variable(&self.var_name, VBValue::String(rhs_str.into()));
                        return Ok(());
                    }
                }
            }
        }
        let value = evaluate(&self.expr, context)?;
        context.set_variable(&self.var_name, value);
        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        compiler.compile_expr(&self.expr);
        let name_lower = self.var_name.to_lowercase();
        if let Some(slot) = compiler.local_slot(&name_lower) {
            compiler.emit(Instruction::StoreLocal(slot));
        } else {
            let idx = compiler.add_constant(VBValue::String(name_lower.into()));
            compiler.emit(Instruction::StoreGlobal(idx));
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
