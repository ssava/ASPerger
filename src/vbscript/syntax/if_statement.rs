use super::VBSyntax;
use crate::vbscript::{ExecutionContext, VBScriptInterpreter, VBValue};

pub struct IfStatement {
    condition: String,
    then_code: String,
}

impl IfStatement {
    pub fn new(condition: String, then_code: String) -> Self {
        IfStatement { condition, then_code }
    }

    fn evaluate_condition(&self, context: &mut ExecutionContext) -> Result<bool, String> {
        let condition_parts: Vec<&str> = self.condition.split_whitespace().collect();
        if condition_parts.len() != 3 {
            return Err("Condizione non valida".to_string());
        }

        let left = if let Some(value) = context.get_variable(condition_parts[0]) {
            value
        } else if let Ok(num) = condition_parts[0].parse::<f64>() {
            VBValue::Number(num)
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        let operator = condition_parts[1];
        let right = if condition_parts[2].starts_with('"') {
            VBValue::String(condition_parts[2].trim_matches('"').to_string())
        } else if let Ok(num) = condition_parts[2].parse::<f64>() {
            VBValue::Number(num)
        } else if let Some(value) = context.get_variable(condition_parts[2]) {
            value
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        match operator {
            "=" => Ok(left == right),
            ">" => match (left, right) {
                (VBValue::Number(a), VBValue::Number(b)) => Ok(a > b),
                _ => Ok(false),
            },
            "<" => match (left, right) {
                (VBValue::Number(a), VBValue::Number(b)) => Ok(a < b),
                _ => Ok(false),
            },
            _ => Err("Operatore non supportato".to_string()),
        }
    }
}

impl VBSyntax for IfStatement {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), String> {
        if self.evaluate_condition(context)? {
            let interpreter = VBScriptInterpreter;
            interpreter.execute(&self.then_code, context)?;
        }
        Ok(())
    }
}