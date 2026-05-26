use crate::vbscript::{vbs_error::{VBSError, VBSErrorType}, ExecutionContext, VBValue};
use super::VBSyntax;

pub struct Assignment {
    var_name: String,
    value: String,
}

impl Assignment {
    pub fn new(var_name: String, value: String) -> Self {
        Assignment { var_name, value }
    }
}

impl VBSyntax for Assignment {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let value = self.value.trim();

        if let Some(var_value) = context.get_variable(value) {
            context.set_variable(&self.var_name, var_value);
            return Ok(());
        }

        match value.parse::<VBValue>() {
            Ok(vb_val) => {
                context.set_variable(&self.var_name, vb_val);
                Ok(())
            }
            Err(_) => Err(VBSErrorType::RuntimeError.into_error(format!("Invalid value for assignment: {}", value))),
        }
    }
}