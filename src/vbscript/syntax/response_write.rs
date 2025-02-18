use super::VBSyntax;
use crate::vbscript::{vbs_error::{VBSError, VBSErrorType}, ExecutionContext};

pub struct ResponseWrite {
    content: String,
}

impl ResponseWrite {
    pub fn new(content: String) -> Self {
        ResponseWrite { content }
    }

    fn extract_content(content: &str) -> String {
        // Remove outer parentheses if present
        let content = if content.starts_with('(') && content.ends_with(')') {
            &content[1..content.len() - 1]
        } else {
            content
        };
        content.trim().to_string()
    }
}

impl VBSyntax for ResponseWrite {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let content = Self::extract_content(&self.content);

        // Handle string literals
        if content.starts_with('"') && content.ends_with('"') {
            let content = content.trim_matches('"');
            context.write(content);
            return Ok(());
        }

        // Handle variables
        if let Some(value) = context.get_variable(&content) {
            context.write(&value.to_string());
            return Ok(());
        }

        // Handle numbers
        if let Ok(num) = content.parse::<f64>() {
            context.write(&num.to_string());
            return Ok(());
        }

        // Handle concatenated strings or variables
        if content.contains('&') {
            let parts: Vec<&str> = content.split('&').map(|s| s.trim()).collect();
            let mut result = String::new();
            for part in parts {
                if part.starts_with('"') && part.ends_with('"') {
                    result.push_str(&part[1..part.len() - 1]);
                } else if let Some(value) = context.get_variable(part) {
                    result.push_str(&value.to_string());
                } else if let Ok(num) = part.parse::<f64>() {
                    result.push_str(&num.to_string());
                } else {
                    return Err(VBSErrorType::ValueError.into_error(format!("Valore non valido per Response.Write: {}", part)));
                }
            }
            context.write(&result);
            return Ok(());
        }

        Err(VBSErrorType::ValueError.into_error(format!("Valore non valido per Response.Write: {}", content)))
    }
}