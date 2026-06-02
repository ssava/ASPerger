//! Utility functions for converting `VBValue` instances to primitive types.

use super::value::VBValue;

/// Convert a `VBValue` to its string representation.
pub fn to_arg_string(val: &VBValue) -> String {
    match val {
        VBValue::String(s) => s.clone(),
        VBValue::Null => "Null".to_string(),
        VBValue::Empty => "".to_string(),
        VBValue::Number(n) => n.to_string(),
        VBValue::Boolean(true) => "True".to_string(),
        VBValue::Boolean(false) => "False".to_string(),
        VBValue::Array(_) => "Array".to_string(),
        VBValue::Object(_) => "Object".to_string(),
    }
}

pub fn to_arg_f64(val: &VBValue) -> f64 {
    match val {
        VBValue::Number(n) => *n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty | VBValue::Array(_) | VBValue::Object(_) => 0.0,
    }
}

pub fn to_boolean(val: &VBValue) -> bool {
    match val {
        VBValue::Boolean(b) => *b,
        VBValue::Number(n) => *n != 0.0,
        VBValue::String(s) => !s.is_empty() && s.to_uppercase() != "FALSE" && s != "0",
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(_) | VBValue::Object(_) => true,
    }
}
