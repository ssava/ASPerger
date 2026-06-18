//! Core VBScript value type (`VBValue`) representing all script-level
//! data: strings, numbers, booleans, null, empty, arrays, and objects.

use super::vbobject::VBScriptObject;
use std::sync::Arc;
use std::{fmt, str::FromStr};

#[derive(Debug)]
pub enum VBValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Null,
    Empty,
    #[allow(dead_code)]
    Array(Arc<Vec<VBValue>>, Vec<usize>),
    Object(Box<dyn VBScriptObject>),
}

impl Clone for VBValue {
    fn clone(&self) -> Self {
        match self {
            VBValue::String(s) => VBValue::String(s.clone()),
            VBValue::Number(n) => VBValue::Number(*n),
            VBValue::Boolean(b) => VBValue::Boolean(*b),
            VBValue::Null => VBValue::Null,
            VBValue::Empty => VBValue::Empty,
            VBValue::Array(v, dims) => VBValue::Array(Arc::clone(v), dims.clone()),
            VBValue::Object(obj) => VBValue::Object(obj.clone_box()),
        }
    }
}

impl PartialEq for VBValue {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (VBValue::String(a), VBValue::String(b)) => a == b,
            (VBValue::Number(a), VBValue::Number(b)) => (a - b).abs() < f64::EPSILON,
            (VBValue::Boolean(a), VBValue::Boolean(b)) => a == b,
            (VBValue::Null, VBValue::Null) => true,
            (VBValue::Empty, VBValue::Empty) => true,
            (VBValue::Array(a, _), VBValue::Array(b, _)) => a == b,
            (VBValue::Object(_), VBValue::Object(_)) => false,
            _ => false,
        }
    }
}

impl fmt::Display for VBValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            VBValue::String(s) => write!(f, "{}", s),
            VBValue::Number(n) => write!(f, "{}", n),
            VBValue::Boolean(b) => {
                if *b {
                    write!(f, "True")
                } else {
                    write!(f, "False")
                }
            }
            VBValue::Null => write!(f, "null"),
            VBValue::Empty => write!(f, "Empty"),
            VBValue::Array(v, _) => write!(f, "Array({})", v.len()),
            VBValue::Object(_) => write!(f, "Object"),
        }
    }
}

/// Implements the `FromStr` trait to parse a string into a `VBValue`.
impl FromStr for VBValue {
    type Err = String;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        // Trim whitespace and handle different types
        let trimmed = s.trim();

        if trimmed.eq_ignore_ascii_case("true") {
            Ok(VBValue::Boolean(true))
        } else if trimmed.eq_ignore_ascii_case("false") {
            Ok(VBValue::Boolean(false))
        } else if trimmed.eq_ignore_ascii_case("null") {
            Ok(VBValue::Null)
        } else if let Ok(num) = trimmed.parse::<f64>() {
            // Parse as a number if possible
            Ok(VBValue::Number(num))
        } else if trimmed.starts_with('"') && trimmed.ends_with('"') {
            // Parse as a string (remove surrounding quotes)
            Ok(VBValue::String(trimmed[1..trimmed.len() - 1].to_string()))
        } else {
            Err(format!("Cannot interpret '{}' as a valid value", s))
        }
    }
}
