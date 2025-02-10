use super::VBSyntax;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBValue};

pub struct Function {
    name: String,
    params: Vec<String>,
    body: String,
}

impl Function {
    pub fn new(name: String, params: Vec<String>, body: String) -> Self {
        Function { name, params, body }
    }
}

impl VBSyntax for Function {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        // Store the function definition in the context
        context.set_variable(&self.name, VBValue::Function(self.params.clone(), self.body.clone()));
        Ok(())
    }
}