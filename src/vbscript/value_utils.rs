//! Utility functions for converting `VBValue` instances to primitive types.

use super::value::VBValue;

/// Convert a `VBValue` to its string representation.
pub fn to_arg_string(val: &VBValue) -> String {
    match val {
        VBValue::String(s) => s.to_string(),
        VBValue::Null => "Null".to_string(),
        VBValue::Empty => "".to_string(),
        VBValue::Number(n) => n.to_string(),
        VBValue::Boolean(true) => "True".to_string(),
        VBValue::Boolean(false) => "False".to_string(),
        VBValue::Array(..) => "Array".to_string(),
        VBValue::Object(_) => "Object".to_string(),
    }
}

pub fn to_arg_f64(val: &VBValue) -> f64 {
    match val {
        VBValue::Number(n) => *n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty | VBValue::Array(..) | VBValue::Object(_) => 0.0,
    }
}

/// Compute the flat index into a row-major array given per-dimension indices.
/// Returns `None` if the index count mismatches or any index is out of range.
pub fn compute_flat_index(indices: &[VBValue], dims: &[usize]) -> Option<usize> {
    if indices.len() != dims.len() {
        return None;
    }
    let mut idx = 0usize;
    for (i, dim) in dims.iter().enumerate() {
        let d = to_arg_f64(&indices[i]) as usize;
        if d > *dim {
            return None;
        }
        idx = idx * (dim + 1) + d;
    }
    Some(idx)
}

pub fn to_boolean(val: &VBValue) -> bool {
    match val {
        VBValue::Boolean(b) => *b,
        VBValue::Number(n) => *n != 0.0,
        VBValue::String(s) => !s.is_empty() && !s.eq_ignore_ascii_case("false") && s.as_ref() != "0",
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(..) | VBValue::Object(_) => true,
    }
}
