//! VBScript.RegExp COM object implementation for regular expression matching
//! and replacement.

use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::value_utils;
use super::vbobject::VBScriptObject;
use super::vbs_error::{VBSError, VBSErrorType};
use crate::{impl_vbscript_object, prop_not_found, method_not_found, cannot_set_property};
use regex::Regex;
use regex::Match as RegexMatch;

/// `VBScript.RegExp` — regular expression matching and replacement.
///
/// Properties: `Pattern`, `IgnoreCase`, `Global`.
/// Methods: `Test(string)` → Boolean, `Execute(string)` → Matches,
/// `Replace(string, replacement)` → String.
#[derive(Debug, Clone)]
pub struct RegExpObject {
    pattern: String,
    ignore_case: bool,
    global: bool,
}

impl Default for RegExpObject {
    fn default() -> Self {
        Self::new()
    }
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
    impl_vbscript_object!(RegExpObject, "RegExp");

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "PATTERN" => Ok(VBValue::String(self.pattern.clone())),
            "IGNORECASE" => Ok(VBValue::Boolean(self.ignore_case)),
            "GLOBAL" => Ok(VBValue::Boolean(self.global)),
            _ => prop_not_found!("RegExp", name),
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
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
            _ => cannot_set_property!("RegExp", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "TEST" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("RegExp.Test requires at least 1 argument".to_string()));
                }
                let input = value_utils::to_arg_string(&args[0]);
                let re = self.compile()?;
                Ok(VBValue::Boolean(re.is_match(&input)))
            }
            "EXECUTE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("RegExp.Execute requires at least 1 argument".to_string()));
                }
                let input = value_utils::to_arg_string(&args[0]);
                let re = self.compile()?;
                let matches: Vec<VBValue> = if self.global {
                    re.captures_iter(&input)
                        .map(|c| {
                            let m = c.get(0).unwrap();
                            let subs: Vec<String> = c.iter().skip(1)
                                .map(|opt| opt.map(|m| m.as_str().to_string()).unwrap_or_default())
                                .collect();
                            VBValue::Object(Box::new(MatchObject::with_submatches(m, &input, subs)))
                        })
                        .collect()
                } else {
                    re.captures(&input)
                        .map(|c| {
                            let m = c.get(0).unwrap();
                            let subs: Vec<String> = c.iter().skip(1)
                                .map(|opt| opt.map(|m| m.as_str().to_string()).unwrap_or_default())
                                .collect();
                            vec![VBValue::Object(Box::new(MatchObject::with_submatches(m, &input, subs)))]
                        })
                        .unwrap_or_default()
                };
                Ok(VBValue::Array(std::sync::Arc::new(matches), vec![]))
            }
            "REPLACE" => {
                if args.len() < 2 {
                    return Err(VBSErrorType::ValueError
                        .into_error("RegExp.Replace requires at least 2 arguments".to_string()));
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
            _ => method_not_found!("RegExp", name),
        }
    }
}

// ===== MatchObject =====

#[derive(Debug, Clone)]
/// A single `RegExp` match result.  Properties: `Value`, `FirstIndex`, `Length`, `SubMatches`.
pub(crate) struct MatchObject {
    value: String,
    first_index: usize,
    length: usize,
    sub_matches: Vec<String>,
}

impl MatchObject {
    pub fn with_submatches(m: RegexMatch, _input: &str, sub_matches: Vec<String>) -> Self {
        MatchObject {
            value: m.as_str().to_string(),
            first_index: m.start(),
            length: m.len(),
            sub_matches,
        }
    }
}

impl VBScriptObject for MatchObject {
    impl_vbscript_object!(MatchObject, "Match");

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "VALUE" => Ok(VBValue::String(self.value.clone())),
            "FIRSTINDEX" => Ok(VBValue::Number(self.first_index as f64)),
            "LENGTH" => Ok(VBValue::Number(self.length as f64)),
            "SUBMATCHES" => Ok(VBValue::Object(Box::new(SubMatchesObject::new(self.sub_matches.clone())))),
            _ => prop_not_found!("Match", name),
        }
    }

    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let i = value_utils::to_arg_f64(index) as usize;
        self.sub_matches.get(i).cloned()
            .map(VBValue::String)
            .ok_or_else(|| VBSErrorType::RuntimeError
                .into_error(format!("SubMatch index {} out of range", i)))
    }

    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

// ===== SubMatchesObject =====

#[derive(Debug, Clone)]
/// Collection of `RegExp` submatch strings, indexed by position.
pub(crate) struct SubMatchesObject {
    items: Vec<String>,
}

impl SubMatchesObject {
    pub fn new(items: Vec<String>) -> Self {
        SubMatchesObject { items }
    }
}

impl VBScriptObject for SubMatchesObject {
    impl_vbscript_object!(SubMatchesObject, "SubMatches");

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.items.len() as f64)),
            _ => prop_not_found!("SubMatches", name),
        }
    }

    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let i = value_utils::to_arg_f64(index) as usize;
        self.items.get(i).cloned()
            .map(VBValue::String)
            .ok_or_else(|| VBSErrorType::RuntimeError
                .into_error(format!("SubMatches index {} out of range", i)))
    }

    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}
