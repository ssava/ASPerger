use std::{fmt, str::FromStr};

#[derive(Clone, Debug, PartialEq)]
pub enum VBValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Null,
}

impl fmt::Display for VBValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            VBValue::String(s) => write!(f, "{}", s),
            VBValue::Number(n) => write!(f, "{}", n),
            VBValue::Boolean(b) => write!(f, "{}", b),
            VBValue::Null => write!(f, "null"),
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