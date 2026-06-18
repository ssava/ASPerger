use super::VBSyntax;
use crate::vbscript::expr::{evaluate, BinOp, Expr};
use crate::vbscript::value_utils;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

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
                                s.push_str(&rhs_str);
                                return Ok(());
                            }
                            VBValue::Empty => {
                                *existing = VBValue::String(rhs_str);
                                return Ok(());
                            }
                            _ => {
                                let lhs_str = existing.to_string();
                                *existing = VBValue::String(lhs_str + &rhs_str);
                                return Ok(());
                            }
                        }
                    } else {
                        context.set_variable(&self.var_name, VBValue::String(rhs_str));
                        return Ok(());
                    }
                }
            }
        }
        let value = evaluate(&self.expr, context)?;
        context.set_variable(&self.var_name, value);
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
