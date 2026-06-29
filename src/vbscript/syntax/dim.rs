use super::VBSyntax;
use crate::vbscript::compiler::Compiler;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::instruction::Instruction;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

/// AST node for `Dim var` / `Dim var(5)` / `Dim var(2, 3)` declarations.
///
/// Each entry in `var_names` is a `(name, optional_dimension_bounds)` pair.
/// When dimension bounds are given, an `Array` of `Empty` values is allocated
/// with the corresponding size.  Without bounds, the variable is set to `Empty`.
#[derive(Clone)]
pub struct Dim {
    var_names: Vec<(String, Option<Vec<Expr>>)>,
}

impl Dim {
    pub fn new(var_names: Vec<(String, Option<Vec<Expr>>)>) -> Self {
        Dim { var_names }
    }
}

impl VBSyntax for Dim {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        for (var_name, dims) in &self.var_names {
            match dims {
                None => {
                    context.set_variable(var_name, VBValue::Empty);
                }
                Some(dim_exprs) if dim_exprs.is_empty() => {
                    context.set_variable(var_name, VBValue::Array(std::sync::Arc::new(Vec::new()), vec![]));
                }
                Some(dim_exprs) => {
                    let mut dim_bounds = Vec::new();
                    let mut total_size = 1usize;
                    for dim_expr in dim_exprs {
                        let val = evaluate(dim_expr, context)?;
                        let bound = to_number(&val) as usize;
                        dim_bounds.push(bound);
                        total_size *= bound + 1;
                    }
                    context.set_variable(
                        var_name,
                        VBValue::Array(
                            std::sync::Arc::new(vec![VBValue::Empty; total_size]),
                            dim_bounds,
                        ),
                    );
                }
            }
        }
        Ok(())
    }

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        for (var_name, dims) in &self.var_names {
            let slot = compiler.allocate_local(var_name);
            match dims {
                None => {
                    compiler.emit(Instruction::LoadEmpty);
                    compiler.emit(Instruction::StoreLocal(slot));
                }
                Some(dim_exprs) if dim_exprs.is_empty() => {
                    compiler.emit(Instruction::NewArray(0));
                    compiler.emit(Instruction::StoreLocal(slot));
                }
                Some(dim_exprs) => {
                    for dim_expr in dim_exprs {
                        compiler.compile_expr(dim_expr);
                    }
                    compiler.emit(Instruction::NewArray(dim_exprs.len() as u8));
                    compiler.emit(Instruction::StoreLocal(slot));
                }
            }
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
