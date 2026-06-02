use std::sync::Arc;
use super::VBSyntax;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::value::VBValue;
use crate::vbscript::{vbs_error::VBSError, vbs_error::VBSErrorType, ExecutionContext};

pub struct ArrayAssignment {
    var_name: String,
    index_expr: Expr,
    value_expr: Expr,
}

impl ArrayAssignment {
    pub fn new(var_name: String, index_expr: Expr, value_expr: Expr) -> Self {
        ArrayAssignment { var_name, index_expr, value_expr }
    }
}

impl VBSyntax for ArrayAssignment {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let idx_val = evaluate(&self.index_expr, context)?;
        let value = evaluate(&self.value_expr, context)?;

        // Check type first to avoid borrow conflicts when swapping object out
        let is_object = matches!(context.get_variable(&self.var_name), Some(VBValue::Object(_)));
        let is_array = matches!(context.get_variable(&self.var_name), Some(VBValue::Array(_)));

        if is_object {
            let mut obj_val = {
                let slot = match context.get_variable_mut(&self.var_name) {
                    Some(v @ VBValue::Object(_)) => v,
                    _ => unreachable!(),
                };
                let mut replacement = VBValue::Empty;
                std::mem::swap(slot, &mut replacement);
                replacement
            };
            match &mut obj_val {
                VBValue::Object(ref mut obj) => {
                    obj.indexed_set(&idx_val, value, context)?;
                }
                _ => unreachable!(),
            }
            context.set_variable(&self.var_name.to_uppercase(), obj_val);
            Ok(())
        } else if is_array {
            match context.get_variable_mut(&self.var_name) {
                Some(VBValue::Array(ref mut items)) => {
                    let idx = to_number(&idx_val) as usize;
                    let items = Arc::make_mut(items);
                    if idx >= items.len() {
                        return Err(VBSErrorType::RuntimeError.into_error(
                            format!("Subscript out of range: index {} exceeds array size {}", idx, items.len())
                        ));
                    }
                    items[idx] = value;
                    Ok(())
                }
                _ => unreachable!(),
            }
        } else {
            match context.get_variable(&self.var_name) {
                Some(_) => Err(VBSErrorType::ValueError.into_error(
                    format!("Variable '{}' does not support indexed assignment", self.var_name)
                )),
                None => Err(VBSErrorType::RuntimeError.into_error(
                    format!("Variable '{}' is not defined", self.var_name)
                )),
            }
        }
    }
}
