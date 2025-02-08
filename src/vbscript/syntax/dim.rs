use crate::vbscript::{ExecutionContext, VBValue};
use super::VBSyntax;

pub struct Dim {
    var_names: Vec<String>,
}

impl Dim {
    pub fn new(var_names: Vec<String>) -> Self {
        Dim { var_names }
    }
}

impl VBSyntax for Dim {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), String> {
        for var_name in &self.var_names {
            context.set_variable(var_name, VBValue::Null);
        }
        Ok(())
    }
}