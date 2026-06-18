use super::VBSyntax;
use crate::vbscript::expr::{evaluate, to_number, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

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

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
