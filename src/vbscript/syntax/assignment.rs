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

        if value.starts_with('"') && value.ends_with('"') {
            let string_value = value.trim_matches('"').to_string();
            context.set_variable(&self.var_name, VBValue::String(string_value));
            return Ok(());
        }

        if let Ok(num) = value.parse::<f64>() {
            context.set_variable(&self.var_name, VBValue::Number(num));
            return Ok(());
        }

        match value.to_lowercase().as_str() {
            "true" => {
                context.set_variable(&self.var_name, VBValue::Boolean(true));
                return Ok(());
            }
            "false" => {
                context.set_variable(&self.var_name, VBValue::Boolean(false));
                return Ok(());
            }
            _ => {}
        }

        if let Some(var_value) = context.get_variable(value) {
            context.set_variable(&self.var_name, var_value);
            return Ok(());
        }

        Err(VBSErrorType::RuntimeError.into_error(format!("Valore non valido per l'assegnazione: {}", value)))
    }
}