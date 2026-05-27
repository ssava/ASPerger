use super::VBSyntax;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::value::VBValue;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

pub struct ReDim {
    var_name: String,
    size_expr: Expr,
    preserve: bool,
}

impl ReDim {
    pub fn new(var_name: String, size_expr: Expr, preserve: bool) -> Self {
        ReDim { var_name, size_expr, preserve }
    }
}

impl VBSyntax for ReDim {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let size_val = evaluate(&self.size_expr, context)?;
        let size = to_number(&size_val) as usize;
        let new_len = size + 1;

        if self.preserve {
            match context.get_variable(&self.var_name) {
                Some(&VBValue::Array(ref old_items)) => {
                    let mut items = vec![VBValue::Empty; new_len];
                    let copy_len = old_items.len().min(new_len);
                    items[..copy_len].clone_from_slice(&old_items[..copy_len]);
                    context.set_variable(&self.var_name, VBValue::Array(items));
                }
                _ => {
                    context.set_variable(&self.var_name, VBValue::Array(vec![VBValue::Empty; new_len]));
                }
            }
        } else {
            context.set_variable(&self.var_name, VBValue::Array(vec![VBValue::Empty; new_len]));
        }

        Ok(())
    }
}
