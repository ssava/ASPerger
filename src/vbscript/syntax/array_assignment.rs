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
        let idx = to_number(&idx_val) as usize;
        let value = evaluate(&self.value_expr, context)?;

        match context.get_variable_mut(&self.var_name) {
            Some(VBValue::Array(ref mut items)) => {
                let items = std::sync::Arc::make_mut(items);
                if idx >= items.len() {
                    return Err(VBSErrorType::RuntimeError.into_error(
                        format!("Subscript out of range: index {} exceeds array size {}", idx, items.len())
                    ));
                }
                items[idx] = value;
                Ok(())
            }
            _ => Err(VBSErrorType::ValueError.into_error(
                format!("Variable '{}' is not an array", self.var_name)
            )),
        }
    }
}
