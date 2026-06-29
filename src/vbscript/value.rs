//! Core VBScript value type (`VBValue`) representing all script-level
//! data: strings, numbers, booleans, null, empty, arrays, and objects.

use super::vbobject::VBScriptObject;
use std::sync::Arc;
use std::fmt;

#[derive(Debug)]
pub enum VBValue {
    String(Arc<str>),
    Number(f64),
    Boolean(bool),
    Null,
    Empty,
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::str::FromStr;

    impl FromStr for VBValue {
        type Err = String;

        fn from_str(s: &str) -> Result<Self, Self::Err> {
            let trimmed = s.trim();
            if trimmed.eq_ignore_ascii_case("true") {
                Ok(VBValue::Boolean(true))
            } else if trimmed.eq_ignore_ascii_case("false") {
                Ok(VBValue::Boolean(false))
            } else if trimmed.eq_ignore_ascii_case("null") {
                Ok(VBValue::Null)
            } else if let Ok(num) = trimmed.parse::<f64>() {
                Ok(VBValue::Number(num))
            } else if trimmed.starts_with('"') && trimmed.ends_with('"') {
                Ok(VBValue::String(trimmed[1..trimmed.len() - 1].to_string().into()))
            } else {
                Err(format!("Cannot interpret '{}' as a valid value", s))
            }
        }
    }

    #[test]
    fn test_vb_value_clone_string() {
        let v = VBValue::String("hello".into());
        let c = v.clone();
        assert_eq!(v, c);
    }

    #[test]
    fn test_vb_value_clone_number() {
        let v = VBValue::Number(42.5);
        let c = v.clone();
        assert_eq!(v, c);
    }

    #[test]
    fn test_vb_value_clone_boolean() {
        let v = VBValue::Boolean(true);
        let c = v.clone();
        assert_eq!(v, c);
    }

    #[test]
    fn test_vb_value_clone_null() {
        let v = VBValue::Null;
        let c = v.clone();
        assert_eq!(v, c);
    }

    #[test]
    fn test_vb_value_clone_empty() {
        let v = VBValue::Empty;
        let c = v.clone();
        assert_eq!(v, c);
    }

    #[test]
    fn test_vb_value_partial_eq_different_types() {
        assert_ne!(VBValue::Number(1.0), VBValue::String("1".into()));
        assert_ne!(VBValue::Empty, VBValue::Null);
        assert_ne!(VBValue::Boolean(true), VBValue::Number(1.0));
    }

    #[test]
    fn test_vb_value_from_str_true() {
        let v: VBValue = "True".parse().unwrap();
        assert_eq!(v, VBValue::Boolean(true));
    }

    #[test]
    fn test_vb_value_from_str_false() {
        let v: VBValue = "false".parse().unwrap();
        assert_eq!(v, VBValue::Boolean(false));
    }

    #[test]
    fn test_vb_value_from_str_null() {
        let v: VBValue = "Null".parse().unwrap();
        assert_eq!(v, VBValue::Null);
    }

    #[test]
    fn test_vb_value_from_str_number() {
        let v: VBValue = "42".parse().unwrap();
        match v {
            VBValue::Number(n) => assert!((n - 42.0).abs() < 1e-10),
            _ => panic!("expected Number"),
        }
    }

    #[test]
    fn test_vb_value_from_str_string() {
        let v: VBValue = "\"hello\"".parse().unwrap();
        assert_eq!(v, VBValue::String("hello".into()));
    }

    #[test]
    fn test_vb_value_from_str_invalid() {
        let result: Result<VBValue, String> = "not a value".parse();
        assert!(result.is_err());
    }

    #[test]
    fn test_vb_value_display_string() {
        let v = VBValue::String("test".into());
        assert_eq!(v.to_string(), "test");
    }

    #[test]
    fn test_vb_value_display_number() {
        let v = VBValue::Number(42.0);
        assert_eq!(v.to_string(), "42");
    }

    #[test]
    fn test_vb_value_display_boolean_true() {
        let v = VBValue::Boolean(true);
        assert_eq!(v.to_string(), "True");
    }

    #[test]
    fn test_vb_value_display_boolean_false() {
        let v = VBValue::Boolean(false);
        assert_eq!(v.to_string(), "False");
    }

    #[test]
    fn test_vb_value_display_null() {
        let v = VBValue::Null;
        assert_eq!(v.to_string(), "null");
    }

    #[test]
    fn test_vb_value_display_empty() {
        let v = VBValue::Empty;
        assert_eq!(v.to_string(), "Empty");
    }

    #[test]
    fn test_vb_value_display_array() {
        let v = VBValue::Array(Arc::new(vec![VBValue::Number(1.0)]), vec![]);
        assert_eq!(v.to_string(), "Array(1)");
    }

    #[test]
    fn test_vb_value_number_partial_eq() {
        let a = VBValue::Number(1.0);
        let b = VBValue::Number(1.0 + f64::EPSILON / 2.0);
        assert_eq!(a, b);
    }
}
