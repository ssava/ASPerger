use super::VBSyntax;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBScriptInterpreter};

pub struct WhileLoop {
    condition: String,
    body: String,
}

impl WhileLoop {
    pub fn new(condition: String, body: String) -> Self {
        WhileLoop { condition, body }
    }
}

impl VBSyntax for WhileLoop {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let interpreter = VBScriptInterpreter;
        while interpreter.evaluate_condition(&self.condition, context)? {
            interpreter.execute(&self.body, context)?;
        }
        Ok(())
    }
}