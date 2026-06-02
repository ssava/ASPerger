use regex::Regex;
use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::value_utils;
use super::vbs_error::{VBSError, VBSErrorType};
use super::vbobject::VBScriptObject;

#[derive(Debug, Clone)]
pub struct RegExpObject {
    pattern: String,
    ignore_case: bool,
    global: bool,
}

impl RegExpObject {
    pub fn new() -> Self {
        RegExpObject {
            pattern: String::new(),
            ignore_case: false,
            global: false,
        }
    }

    fn compile(&self) -> Result<Regex, VBSError> {
        let mut p = self.pattern.clone();
        if self.ignore_case {
            p = format!("(?i){}", p);
        }
        Regex::new(&p).map_err(|e| {
            VBSErrorType::RuntimeError.into_error(format!("Invalid regex pattern: {}", e))
        })
    }
}

impl VBScriptObject for RegExpObject {
    fn type_name(&self) -> &'static str { "RegExp" }

    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "PATTERN" => Ok(VBValue::String(self.pattern.clone())),
            "IGNORECASE" => Ok(VBValue::Boolean(self.ignore_case)),
            "GLOBAL" => Ok(VBValue::Boolean(self.global)),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Property '{}' not found on RegExp object", name)
            )),
        }
    }

    fn set_property(&mut self, name: &str, value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            "PATTERN" => {
                self.pattern = value_utils::to_arg_string(&value);
                Ok(())
            }
            "IGNORECASE" => {
                self.ignore_case = value_utils::to_boolean(&value);
                Ok(())
            }
            "GLOBAL" => {
                self.global = value_utils::to_boolean(&value);
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Cannot set property '{}' on RegExp object", name)
            )),
        }
    }

    fn call_method(&mut self, name: &str, args: &[VBValue], _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "TEST" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "RegExp.Test requires at least 1 argument".to_string()
                    ));
                }
                let input = value_utils::to_arg_string(&args[0]);
                let re = self.compile()?;
                Ok(VBValue::Boolean(re.is_match(&input)))
            }
            "EXECUTE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "RegExp.Execute requires at least 1 argument".to_string()
                    ));
                }
                let input = value_utils::to_arg_string(&args[0]);
                let re = self.compile()?;
                let matches: Vec<VBValue> = if self.global {
                    re.find_iter(&input).map(|m| {
                        VBValue::String(m.as_str().to_string())
                    }).collect()
                } else {
                    re.find(&input).map(|m| {
                        vec![VBValue::String(m.as_str().to_string())]
                    }).unwrap_or_default()
                };
                Ok(VBValue::Array(std::sync::Arc::new(matches)))
            }
            "REPLACE" => {
                if args.len() < 2 {
                    return Err(VBSErrorType::ValueError.into_error(
                        "RegExp.Replace requires at least 2 arguments".to_string()
                    ));
                }
                let input = value_utils::to_arg_string(&args[0]);
                let replacement = value_utils::to_arg_string(&args[1]);
                let re = self.compile()?;
                let result = if self.global {
                    re.replace_all(&input, replacement.as_str())
                } else {
                    re.replace(&input, replacement.as_str())
                };
                Ok(VBValue::String(result.to_string()))
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Method '{}' not found on RegExp object", name)
            )),
        }
    }
}
