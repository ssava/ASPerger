use regex::Regex;
use super::{ExecutionContext, VBValue};

pub struct VBScriptInterpreter;

impl VBScriptInterpreter {
    pub fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let code = code.trim();

        for line in code.split('\n') {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }

            if line.starts_with('\'') || line.to_lowercase().starts_with("rem") {
                continue;
            }

            if line.to_lowercase().starts_with("response.write") {
                self.handle_response_write(line, context)?;
            } else if line.to_lowercase().starts_with("dim") {
                self.handle_dim(line, context)?;
            } else if line.contains('=') {
                self.handle_assignment(line, context)?;
            } else if line.to_lowercase().starts_with("if") {
                self.handle_if(line, context)?;
            }
        }
        Ok(())
    }

    fn handle_response_write(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let content = code
            .trim()
            .strip_prefix("Response.Write")
            .unwrap_or("")
            .trim();

        let content = if content.starts_with('(') && content.ends_with(')') {
            &content[1..content.len()-1]
        } else {
            content
        };

        let content = content.trim();

        if content.starts_with('"') && content.ends_with('"') {
            let content = &content[1..content.len()-1];
            context.write(content);
            return Ok(());
        }

        if let Some(value) = context.get_variable(content) {
            context.write(&value.to_string());
            return Ok(());
        }

        if let Ok(num) = content.parse::<f64>() {
            context.write(&num.to_string());
            return Ok(());
        }

        Err(format!("Valore non valido per Response.Write: {}", content))
    }

    fn handle_dim(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let var_names = code
            .trim()
            .strip_prefix("Dim")
            .unwrap_or("")
            .split(',')
            .map(|s| s.trim())
            .filter(|s| !s.is_empty());

        for var_name in var_names {
            context.set_variable(var_name, VBValue::Null);
        }
        
        Ok(())
    }

    fn handle_assignment(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let parts: Vec<&str> = code.split('=').collect();
        if parts.len() != 2 {
            return Err("Assegnazione non valida".to_string());
        }

        let var_name = parts[0].trim();
        let value = parts[1].trim();

        if value.starts_with('"') && value.ends_with('"') {
            let string_value = value.trim_matches('"').to_string();
            context.set_variable(var_name, VBValue::String(string_value));
            return Ok(());
        }

        if let Ok(num) = value.parse::<f64>() {
            context.set_variable(var_name, VBValue::Number(num));
            return Ok(());
        }

        match value.to_lowercase().as_str() {
            "true" => {
                context.set_variable(var_name, VBValue::Boolean(true));
                return Ok(());
            }
            "false" => {
                context.set_variable(var_name, VBValue::Boolean(false));
                return Ok(());
            }
            _ => {}
        }

        if let Some(var_value) = context.get_variable(value) {
            context.set_variable(var_name, var_value);
            return Ok(());
        }

        Err(format!("Valore non valido per l'assegnazione: {}", value))
    }

    fn handle_if(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let if_pattern = Regex::new(r"If\s+(.+?)\s+Then\s+(.+?)\s+End If").unwrap();
        
        if let Some(caps) = if_pattern.captures(code) {
            let condition = caps.get(1).unwrap().as_str();
            let then_code = caps.get(2).unwrap().as_str();

            if self.evaluate_condition(condition, context)? {
                self.execute(then_code, context)?;
            }
            return Ok(());
        }

        Err("Sintassi If non valida".to_string())
    }

    fn evaluate_condition(&self, condition: &str, context: &mut ExecutionContext) -> Result<bool, String> {
        let parts: Vec<&str> = condition.split_whitespace().collect();
        if parts.len() != 3 {
            return Err("Condizione non valida".to_string());
        }

        let left = if let Some(value) = context.get_variable(parts[0]) {
            value
        } else if let Ok(num) = parts[0].parse::<f64>() {
            VBValue::Number(num)
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        let operator = parts[1];
        let right = if parts[2].starts_with('"') {
            VBValue::String(parts[2].trim_matches('"').to_string())
        } else if let Ok(num) = parts[2].parse::<f64>() {
            VBValue::Number(num)
        } else if let Some(value) = context.get_variable(parts[2]) {
            value
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        match operator {
            "=" => Ok(self.compare_values(&left, &right)),
            ">" => Ok(self.compare_greater(&left, &right)),
            "<" => Ok(self.compare_less(&left, &right)),
            _ => Err("Operatore non supportato".to_string()),
        }
    }

    fn compare_values(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::String(a), VBValue::String(b)) => a == b,
            (VBValue::Number(a), VBValue::Number(b)) => (a - b).abs() < f64::EPSILON,
            (VBValue::Boolean(a), VBValue::Boolean(b)) => a == b,
            _ => false,
        }
    }

    fn compare_greater(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::Number(a), VBValue::Number(b)) => a > b,
            _ => false,
        }
    }

    fn compare_less(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::Number(a), VBValue::Number(b)) => a < b,
            _ => false,
        }
    }
}