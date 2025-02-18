use super::VBSyntax;
use crate::vbscript::{vbs_error::{VBSError, VBSErrorType}, ExecutionContext, VBScriptInterpreter, VBValue};

pub struct CallFunction {
    name: String,
    args: Vec<String>,
}

impl CallFunction {
    pub fn new(name: String, args: Vec<String>) -> Self {
        CallFunction { name, args }
    }
}

impl VBSyntax for CallFunction {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        if let Some(VBValue::Function(params, body)) = context.get_variable(&self.name) {
            if params.len() != self.args.len() {
                return Err(VBSErrorType::SyntaxError.into_error(format!(
                    "Numero di argomenti non valido per la funzione {}",
                    self.name
                )));
            }

            // Set arguments as variables in the context
            for (param, arg) in params.iter().zip(self.args.iter()) {
                let value = if let Some(var_value) = context.get_variable(arg) {
                    var_value
                } else if let Ok(num) = arg.parse::<f64>() {
                    VBValue::Number(num)
                } else if arg.starts_with('"') && arg.ends_with('"') {
                    VBValue::String(arg[1..arg.len() - 1].to_string())
                } else {
                    return Err(VBSErrorType::SyntaxError.into_error(format!("Argomento non valido: {}", arg)));
                };
                context.set_variable(param, value);
            }

            // Execute the function body
            let interpreter = VBScriptInterpreter;
            interpreter.execute(&body, context)?;
        } else {
            return Err(VBSErrorType::RuntimeError.into_error(format!("Funzione non definita: {}", self.name)));
        }
        Ok(())
    }
}