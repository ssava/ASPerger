use super::VBSyntax;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::value::VBValue;
use crate::vbscript::{vbs_error::VBSError, vbs_error::VBSErrorType, ExecutionContext};
use std::sync::Arc;

/// Convert multi-dimensional array indices into a flat offset.
///
/// Each dimension `dims[i]` stores the upper bound; the total range for
/// that dimension is `dims[i] + 1`.  The flat index is computed as:
///   idx = (((i0 * (d1+1)) + i1) * (d2+1)) + i2 ...
/// Returns `None` if any index exceeds its dimension bound.
fn compute_flat_index(indices: &[VBValue], dims: &[usize]) -> Option<usize> {
    if indices.len() != dims.len() {
        return None;
    }
    let mut idx = 0usize;
    for (i, dim) in dims.iter().enumerate() {
        let d = to_number(&indices[i]) as usize;
        if d > *dim {
            return None;
        }
        idx = idx * (dim + 1) + d;
    }
    Some(idx)
}

/// AST node for `arr(i) = value` or `arr(i, j) = value` (array element assignment).
///
/// Supports both plain VBScript arrays and Object-indexed assignment
/// (e.g. `Application("key") = value`), dispatching to `indexed_set`
/// when the target is an `Object` or direct element mutation for `Array`.
#[derive(Clone)]
pub struct ArrayAssignment {
    var_name: String,
    index_exprs: Vec<Expr>,
    value_expr: Expr,
}

impl ArrayAssignment {
    pub fn new(var_name: String, index_exprs: Vec<Expr>, value_expr: Expr) -> Self {
        ArrayAssignment {
            var_name,
            index_exprs,
            value_expr,
        }
    }
}

impl VBSyntax for ArrayAssignment {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let indices: Result<Vec<VBValue>, VBSError> =
            self.index_exprs.iter().map(|e| evaluate(e, context)).collect();
        let indices = indices?;
        let value = evaluate(&self.value_expr, context)?;

        // Check type first to avoid borrow conflicts when swapping object out
        let is_object = matches!(
            context.get_variable(&self.var_name),
            Some(VBValue::Object(_))
        );
        let is_array = matches!(
            context.get_variable(&self.var_name),
            Some(VBValue::Array(..))
        );

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
                    obj.indexed_set(&indices[0], value, context)?;
                }
                _ => unreachable!(),
            }
            context.set_variable(&self.var_name.to_uppercase(), obj_val);
            Ok(())
        } else if is_array {
            match context.get_variable_mut(&self.var_name) {
                Some(VBValue::Array(ref mut items, ref dims)) => {
                    let flat_idx = if dims.is_empty() {
                        // Dynamic array — use first index directly
                        let idx = to_number(&indices[0]) as usize;
                        if idx >= items.len() {
                            return Err(VBSErrorType::RuntimeError.into_error(format!(
                                "Subscript out of range: index {} exceeds array size {}",
                                idx,
                                items.len()
                            )));
                        }
                        idx
                    } else {
                        compute_flat_index(&indices, dims).ok_or_else(|| {
                            VBSErrorType::RuntimeError.into_error(
                                "Subscript out of range".to_string(),
                            )
                        })?
                    };
                    let items = Arc::make_mut(items);
                    if flat_idx >= items.len() {
                        return Err(VBSErrorType::RuntimeError.into_error(format!(
                            "Subscript out of range: index {} exceeds array size {}",
                            flat_idx,
                            items.len()
                        )));
                    }
                    items[flat_idx] = value;
                    Ok(())
                }
                _ => unreachable!(),
            }
        } else {
            match context.get_variable(&self.var_name) {
                Some(_) => Err(VBSErrorType::ValueError.into_error(format!(
                    "Variable '{}' does not support indexed assignment",
                    self.var_name
                ))),
                None => Err(VBSErrorType::RuntimeError
                    .into_error(format!("Variable '{}' is not defined", self.var_name))),
            }
        }
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
