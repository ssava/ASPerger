use std::str::FromStr;

use crate::vbscript::syntax::{
    Assignment, CallFunction, Dim, ForLoop, Function, IfStatement, ResponseWrite, VBSyntax,
    WhileLoop,
};
use crate::vbscript::ExecutionContext;
use regex::Regex;

use super::vbs_error::{VBSError, VBSErrorType};
use super::VBValue;

pub struct VBScriptInterpreter;

impl VBScriptInterpreter {
    /// Executes the provided VBScript code by interpreting each line.
    ///
    /// # Arguments
    /// * `code` - A string slice containing the VBScript code to execute.
    /// * `context` - A mutable reference to the execution context where variables and functions are stored.
    ///
    /// # Returns
    /// * `Ok(())` if the execution is successful.
    /// * `Err(String)` if there is a syntax or runtime error.
    pub fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let code = code.trim();
        for line in code.split('\n') {
            let line = line.trim();
            // Skip empty lines and comments
            if line.is_empty() || line.starts_with('\'') || line.to_lowercase().starts_with("rem") {
                continue;
            }

            // Create a syntax object using the factory method
            match self.create_syntax(line, code)? {
                Some(syntax) => syntax.execute(context)?,
                None => return Err(VBSErrorType::NotImplementedError.into_error(format!("Comando non riconosciuto: {}", line))),
            }
        }
        Ok(())
    }

    /// Factory method to create a VBSyntax object based on the line of code.
    ///
    /// # Arguments
    /// * `line` - A string slice representing a single line of VBScript code.
    /// * `full_code` - A string slice containing the entire VBScript code (used for block extraction).
    ///
    /// # Returns
    /// * `Ok(Some(Box<dyn VBSyntax>))` if the line corresponds to a valid VBScript statement.
    /// * `Ok(None)` if the line does not match any known VBScript statement.
    /// * `Err(VBSError)` if there is a syntax error.
    fn create_syntax(
        &self,
        line: &str,
        full_code: &str,
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        // Handle different types of VBScript statements
        if line.to_lowercase().starts_with("response.write") {
            let content = line
                .trim()
                .strip_prefix("Response.Write")
                .unwrap_or("")
                .trim()
                .to_string();
            Ok(Some(Box::new(ResponseWrite::new(content))))
        } else if line.to_lowercase().starts_with("dim") {
            let var_names = line
                .trim()
                .strip_prefix("Dim")
                .unwrap_or("")
                .split(',')
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .collect();
            Ok(Some(Box::new(Dim::new(var_names))))
        } else if line.contains('=') {
            let parts: Vec<&str> = line.split('=').collect();
            if parts.len() != 2 {
                return Err(
                    VBSErrorType::SyntaxError.into_error("Assegnazione non valida".to_string())
                );
            }
            let var_name = parts[0].trim().to_string();
            let value = parts[1].trim().to_string();
            Ok(Some(Box::new(Assignment::new(var_name, value))))
        } else if line.to_lowercase().starts_with("if") {
            self.parse_if_statement(line)
        } else if line.to_lowercase().starts_with("for") {
            self.parse_for_loop(line, full_code)
        } else if line.to_lowercase().starts_with("while") {
            self.parse_while_loop(line, full_code)
        } else if line.to_lowercase().starts_with("function") {
            self.parse_function(line, full_code)
        } else if line.to_lowercase().starts_with("call") {
            self.parse_call_function(line)
        } else {
            Ok(None) // Unknown command
        }
    }

    /// Evaluates a condition string (e.g., "x > 5") in the given execution context.
    ///
    /// # Arguments
    /// * `condition` - A string slice representing the condition to evaluate.
    /// * `context` - A mutable reference to the execution context where variables are stored.
    ///
    /// # Returns
    /// * `Ok(true)` if the condition evaluates to true.
    /// * `Ok(false)` if the condition evaluates to false.
    /// * `Err(String)` if there is an error in parsing or evaluating the condition.
    pub(crate) fn evaluate_condition(
        &self,
        condition: &str,
        context: &mut ExecutionContext,
    ) -> Result<bool, VBSError> {
        let condition_pattern = Regex::new(r"(\w+)\s*(==|!=|<=|>=|<|>|&|And|Or)\s*(.+?)").unwrap();
        if let Some(caps) = condition_pattern.captures(condition) {
            let lhs_name = caps.get(1).unwrap().as_str();
            let op = caps.get(2).unwrap().as_str();
            let rhs_str = caps.get(3).unwrap().as_str();

            let lhs_value = match context.get_variable(lhs_name) {
                Some(value) => VBValue::from_str(&value.to_string()).map_err(|_| {
                    VBSErrorType::TypeError
                        .into_error(format!("Variabile '{}' non è un tipo valido", lhs_name))
                })?,
                None => {
                    return Err(VBSErrorType::NameError
                        .into_error(format!("Variabile '{}' non definita", lhs_name)))
                }
            };

            let rhs_value = match context.get_variable(rhs_str) {
                Some(value) => VBValue::from_str(&value.to_string()).map_err(|_| {
                    VBSErrorType::TypeError
                        .into_error(format!("Valore destro '{}' non è un tipo valido", rhs_str))
                })?,
                None => VBValue::from_str(rhs_str).map_err(|_| {
                    VBSErrorType::ValueError.into_error(format!(
                        "Impossibile interpretare '{}' come valore",
                        rhs_str
                    ))
                })?,
            };

            match op {
                "==" => Ok(self.compare_values(&lhs_value, &rhs_value)?),
                "!=" => Ok(!self.compare_values(&lhs_value, &rhs_value)?),
                "<=" => self.compare_numeric_values(&lhs_value, &rhs_value, |a, b| a <= b),
                ">=" => self.compare_numeric_values(&lhs_value, &rhs_value, |a, b| a >= b),
                "<" => self.compare_numeric_values(&lhs_value, &rhs_value, |a, b| a < b),
                ">" => self.compare_numeric_values(&lhs_value, &rhs_value, |a, b| a > b),
                "&" => self.combine_strings(&lhs_value, &rhs_value, context, lhs_name),
                "And" => self.evaluate_logical_and(&lhs_value, &rhs_value),
                "Or" => self.evaluate_logical_or(&lhs_value, &rhs_value),
                _ => Err(VBSErrorType::SyntaxError
                    .into_error(format!("Operatore '{}' non supportato", op))),
            }
        } else {
            Err(VBSErrorType::SyntaxError.into_error("Condizione non valida".to_string()))
        }
    }

    /// Compares two `VBValue` instances for equality.
    fn compare_values(&self, lhs: &VBValue, rhs: &VBValue) -> Result<bool, VBSError> {
        match (lhs, rhs) {
            (VBValue::String(l), VBValue::String(r)) => Ok(l == r),
            (VBValue::Number(l), VBValue::Number(r)) => Ok((l - r).abs() < f64::EPSILON), // Handle floating-point precision
            (VBValue::Boolean(l), VBValue::Boolean(r)) => Ok(l == r),
            (VBValue::Null, VBValue::Null) => Ok(true),
            _ => Ok(false), // Different types are always unequal
        }
    }

    /// Compares two numeric `VBValue` instances using a provided comparison function.
    fn compare_numeric_values<F>(
        &self,
        lhs: &VBValue,
        rhs: &VBValue,
        cmp: F,
    ) -> Result<bool, VBSError>
    where
        F: Fn(f64, f64) -> bool,
    {
        match (lhs, rhs) {
            (VBValue::Number(l), VBValue::Number(r)) => Ok(cmp(*l, *r)),
            _ => Err(VBSErrorType::RuntimeError.into_error("Confronto numerico richiede valori di tipo Number".to_string())),
        }
    }

    /// Combines two string `VBValue` instances using the `&` operator.
    fn combine_strings(
        &self,
        lhs: &VBValue,
        rhs: &VBValue,
        context: &mut ExecutionContext,
        lhs_name: &str,
    ) -> Result<bool, VBSError> {
        match (lhs, rhs) {
            (VBValue::String(l), VBValue::String(r)) => {
                // Combine the strings and update the variable in the context
                let combined = format!("{}{}", l, r);
                context.set_variable(lhs_name, VBValue::String(combined));
                Ok(true)
            }
            _ => Err(VBSErrorType::TypeError.into_error("Operatore '&' richiede valori di tipo String".to_string())),
        }
    }

    /// Evaluates the logical AND operation between two `VBValue` instances.
    fn evaluate_logical_and(&self, lhs: &VBValue, rhs: &VBValue) -> Result<bool, VBSError> {
        match (lhs, rhs) {
            (VBValue::Boolean(l), VBValue::Boolean(r)) => Ok(*l && *r),
            _ => Err(VBSErrorType::TypeError.into_error("Operatore 'And' richiede valori di tipo Boolean".to_string())),
        }
    }

    /// Evaluates the logical OR operation between two `VBValue` instances.
    fn evaluate_logical_or(&self, lhs: &VBValue, rhs: &VBValue) -> Result<bool, VBSError> {
        match (lhs, rhs) {
            (VBValue::Boolean(l), VBValue::Boolean(r)) => Ok(*l || *r),
            _ => Err(VBSErrorType::TypeError.into_error("Operatore 'Or' richiede valori di tipo Boolean".to_string())),
        }
    }

    /// Parse an IF statement.
    fn parse_if_statement(&self, line: &str) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let if_pattern = Regex::new(r"If\s+(.+?)\s+Then\s+(.+?)\s+End If").unwrap();
        if let Some(caps) = if_pattern.captures(line) {
            let condition = caps.get(1).unwrap().as_str().to_string();
            let then_code = caps.get(2).unwrap().as_str().to_string();
            Ok(Some(Box::new(IfStatement::new(condition, then_code))))
        } else {
            Err(VBSErrorType::SyntaxError.into_error("Sintassi If non valida".to_string()))
        }
    }

    /// Parse a FOR loop.
    fn parse_for_loop(
        &self,
        line: &str,
        full_code: &str,
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let for_pattern =
            Regex::new(r"For\s+(\w+)\s*=\s*(\d+)\s+To\s+(\d+)(?:\s+Step\s+(\d+))?").unwrap();
        if let Some(caps) = for_pattern.captures(line) {
            let counter = caps.get(1).unwrap().as_str().to_string();
            let start = caps.get(2).unwrap().as_str().parse::<i32>().unwrap();
            let end = caps.get(3).unwrap().as_str().parse::<i32>().unwrap();
            let step = caps
                .get(4)
                .map_or(1, |m| m.as_str().parse::<i32>().unwrap());
            let body = self.extract_body(full_code, "For", "Next")?;
            Ok(Some(Box::new(ForLoop::new(
                counter, start, end, step, body,
            ))))
        } else {
            Err(VBSErrorType::SyntaxError.into_error("Sintassi If non valida".to_string()))
        }
    }

    /// Parse a WHILE loop.
    fn parse_while_loop(
        &self,
        line: &str,
        full_code: &str,
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let condition = line
            .trim()
            .strip_prefix("While")
            .unwrap_or("")
            .trim()
            .to_string();
        let body = self.extract_body(full_code, "While", "Wend")?;
        Ok(Some(Box::new(WhileLoop::new(condition, body))))
    }

    /// Parse a FUNCTION definition.
    fn parse_function(
        &self,
        line: &str,
        full_code: &str,
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let func_pattern = Regex::new(r"Function\s+(\w+)\s*\((.*?)\)").unwrap();
        if let Some(caps) = func_pattern.captures(line) {
            let name = caps.get(1).unwrap().as_str().to_string();
            let params = caps
                .get(2)
                .unwrap()
                .as_str()
                .split(',')
                .map(|s| s.trim().to_string())
                .collect();
            let body = self.extract_body(full_code, "Function", "End Function")?;
            Ok(Some(Box::new(Function::new(name, params, body))))
        } else {
            Err(VBSErrorType::SyntaxError.into_error("Sintassi Function non valida".to_string()))
        }
    }

    /// Parse a CALL function.
    fn parse_call_function(&self, line: &str) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let call_pattern = Regex::new(r"Call\s+(\w+)\s*\((.*?)\)").unwrap();
        if let Some(caps) = call_pattern.captures(line) {
            let name = caps.get(1).unwrap().as_str().to_string();
            let args = caps
                .get(2)
                .unwrap()
                .as_str()
                .split(',')
                .map(|s| s.trim().to_string())
                .collect();
            Ok(Some(Box::new(CallFunction::new(name, args))))
        } else {
            Err(VBSErrorType::BlockMismatchError.into_error("Sintassi Call non valida".to_string()))
        }
    }

    /// Extract the body of a block (e.g., For, While, Function).
    fn extract_body(
        &self,
        code: &str,
        start_keyword: &str,
        end_keyword: &str,
    ) -> Result<String, VBSError> {
        let mut lines = code.lines();
        let mut body = String::new();
        let mut depth = 0; // Track nested blocks

        // Find the start of the block
        while let Some(line) = lines.next() {
            if line
                .trim()
                .to_lowercase()
                .starts_with(&start_keyword.to_lowercase())
            {
                depth += 1;
                break;
            }
        }

        // Extract the body until the end of the block
        while let Some(line) = lines.next() {
            let trimmed_line = line.trim().to_lowercase();
            if trimmed_line.starts_with(&start_keyword.to_lowercase()) {
                depth += 1;
            } else if trimmed_line.starts_with(&end_keyword.to_lowercase()) {
                depth -= 1;
                if depth == 0 {
                    break; // End of the current block
                }
            }
            body.push_str(line);
            body.push('\n');
        }

        if depth != 0 {
            return Err(VBSErrorType::BlockMismatchError.into_error(format!(
                "Blocco non chiuso correttamente: {}...{}", start_keyword, end_keyword)
            ));
        }
        Ok(body)
    }
}
