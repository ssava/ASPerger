use super::VBSyntax;
use crate::vbscript::{vbs_error::VBSError, ExecutionContext, VBScriptInterpreter, VBValue};

pub struct ForLoop {
    counter: String,
    start: i32,
    end: i32,
    step: i32,
    body: String,
}

impl ForLoop {
    pub fn new(counter: String, start: i32, end: i32, step: i32, body: String) -> Self {
        ForLoop {
            counter,
            start,
            end,
            step,
            body,
        }
    }
}

impl VBSyntax for ForLoop {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let mut counter_value = self.start;
        while counter_value <= self.end {
            context.set_variable(&self.counter, VBValue::Number(counter_value as f64));
            let interpreter = VBScriptInterpreter;
            interpreter.execute(&self.body, context)?;
            counter_value += self.step;
        }
        Ok(())
    }
}