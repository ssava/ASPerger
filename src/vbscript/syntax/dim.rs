use super::VBSyntax;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

#[derive(Clone)]
pub struct Dim {
    var_names: Vec<(String, bool)>,
}

impl Dim {
    pub fn new(var_names: Vec<(String, bool)>) -> Self {
        Dim { var_names }
    }
}

impl VBSyntax for Dim {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        for (var_name, is_array) in &self.var_names {
            if *is_array {
                context.set_variable(var_name, VBValue::Array(std::sync::Arc::new(Vec::new())));
            } else {
                context.set_variable(var_name, VBValue::Empty);
            }
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
