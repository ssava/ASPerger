use super::VBSyntax;
use crate::vbscript::compiler::Compiler;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::instruction::Instruction;
use crate::vbscript::value::VBValue;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

/// AST node for `ReDim var(size)` and `ReDim Preserve var(size)`.
///
/// Without `Preserve`, a new array is allocated (content is lost).
/// With `Preserve`, existing elements are copied into the new array up
/// to `min(old_len, new_len)`.  Multi-dimensional `Preserve` is rejected
/// (VBScript limitation — only the last dimension may change).
#[derive(Clone)]
pub struct ReDim {
    var_name: String,
    size_exprs: Vec<Expr>,
    preserve: bool,
}

impl ReDim {
    pub fn new(var_name: String, size_exprs: Vec<Expr>, preserve: bool) -> Self {
        ReDim {
            var_name,
            size_exprs,
            preserve,
        }
    }
}

impl VBSyntax for ReDim {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        // Evaluate all dimension bounds
        let mut dim_bounds = Vec::new();
        let mut total_size = 1usize;
        for expr in &self.size_exprs {
            let val = evaluate(expr, context)?;
            let bound = to_number(&val) as usize;
            dim_bounds.push(bound);
            total_size *= bound + 1;
        }

        if self.preserve {
            if dim_bounds.len() > 1 {
                return Err(crate::vbscript::vbs_error::VBSErrorType::RuntimeError.into_error(
                    "ReDim Preserve can only change the last dimension of a multi-dimensional array"
                        .to_string(),
                ));
            }
            let size = dim_bounds[0];
            let new_len = size + 1;
            match context.get_variable(&self.var_name) {
                Some(VBValue::Array(old_items, _)) => {
                    let mut items = vec![VBValue::Empty; new_len];
                    let copy_len = old_items.len().min(new_len);
                    items[..copy_len].clone_from_slice(&old_items[..copy_len]);
                    context
                        .set_variable(&self.var_name, VBValue::Array(std::sync::Arc::new(items), vec![size]));
                }
                _ => {
                    context.set_variable(
                        &self.var_name,
                        VBValue::Array(std::sync::Arc::new(vec![VBValue::Empty; new_len]), vec![size]),
                    );
                }
            }
        } else {
            context.set_variable(
                &self.var_name,
                VBValue::Array(
                    std::sync::Arc::new(vec![VBValue::Empty; total_size]),
                    dim_bounds,
                ),
            );
        }

        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        let name_lower = self.var_name.to_lowercase();
        let slot = compiler.allocate_local(&name_lower);
        for size_expr in &self.size_exprs {
            compiler.compile_expr(size_expr);
        }
        compiler.emit(Instruction::ReDim(slot, self.size_exprs.len() as u8, self.preserve));
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
