use regex::Regex;
use crate::vbscript::ExecutionContext;
use crate::vbscript::syntax::{VBSyntax, ResponseWrite, Dim, Assignment, IfStatement};

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

            let syntax: Box<dyn VBSyntax> = if line.to_lowercase().starts_with("response.write") {
                let content = line
                    .trim()
                    .strip_prefix("Response.Write")
                    .unwrap_or("")
                    .trim()
                    .to_string();
                Box::new(ResponseWrite::new(content))
            } else if line.to_lowercase().starts_with("dim") {
                let var_names = line
                    .trim()
                    .strip_prefix("Dim")
                    .unwrap_or("")
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();
                Box::new(Dim::new(var_names))
            } else if line.contains('=') {
                let parts: Vec<&str> = line.split('=').collect();
                if parts.len() != 2 {
                    return Err("Assegnazione non valida".to_string());
                }
                let var_name = parts[0].trim().to_string();
                let value = parts[1].trim().to_string();
                Box::new(Assignment::new(var_name, value))
            } else if line.to_lowercase().starts_with("if") {
                let if_pattern = Regex::new(r"If\s+(.+?)\s+Then\s+(.+?)\s+End If").unwrap();
                if let Some(caps) = if_pattern.captures(line) {
                    let condition = caps.get(1).unwrap().as_str().to_string();
                    let then_code = caps.get(2).unwrap().as_str().to_string();
                    Box::new(IfStatement::new(condition, then_code))
                } else {
                    return Err("Sintassi If non valida".to_string());
                }
            } else {
                return Err(format!("Comando non riconosciuto: {}", line));
            };

            syntax.execute(context)?;
        }

        Ok(())
    }
}