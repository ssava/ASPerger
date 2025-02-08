use std::fmt;

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