use super::value::VBValue;
use super::vbobject::Dictionary;
use super::vbs_error::{VBSError, VBSErrorType};

pub fn call_builtin(name: &str, args: Vec<VBValue>) -> Result<VBValue, VBSError> {
    match name {
        n if n.eq_ignore_ascii_case("ARRAY") => builtin_array(&args),
        n if n.eq_ignore_ascii_case("CREATEOBJECT") => builtin_createobject(&args),
        n if n.eq_ignore_ascii_case("LEN") => builtin_len(&args),
        n if n.eq_ignore_ascii_case("UCASE") => builtin_ucase(&args),
        n if n.eq_ignore_ascii_case("LCASE") => builtin_lcase(&args),
        n if n.eq_ignore_ascii_case("MID") => builtin_mid(&args),
        n if n.eq_ignore_ascii_case("LEFT") => builtin_left(&args),
        n if n.eq_ignore_ascii_case("RIGHT") => builtin_right(&args),
        n if n.eq_ignore_ascii_case("TRIM") => builtin_trim(&args),
        n if n.eq_ignore_ascii_case("CINT") => builtin_cint(&args),
        n if n.eq_ignore_ascii_case("CSTR") => builtin_cstr(&args),
        n if n.eq_ignore_ascii_case("ABS") => builtin_abs(&args),
        n if n.eq_ignore_ascii_case("ISNULL") => builtin_isnull(&args),
        n if n.eq_ignore_ascii_case("ISEMPTY") => builtin_isempty(&args),
        n if n.eq_ignore_ascii_case("INSTR") => builtin_instr(&args),
        _ => Err(VBSErrorType::NotImplementedError.into_error(
            format!("Function '{}' is not implemented", name)
        )),
    }
}

fn expect_arg_count(args: &[VBValue], expected: usize, name: &str) -> Result<(), VBSError> {
    if args.len() != expected {
        return Err(VBSErrorType::ValueError.into_error(
            format!("{} requires {} argument(s), got {}", name, expected, args.len())
        ));
    }
    Ok(())
}

fn expect_min_args(args: &[VBValue], min: usize, name: &str) -> Result<(), VBSError> {
    if args.len() < min {
        return Err(VBSErrorType::ValueError.into_error(
            format!("{} requires at least {} argument(s), got {}", name, min, args.len())
        ));
    }
    Ok(())
}

fn builtin_array(args: &[VBValue]) -> Result<VBValue, VBSError> {
    Ok(VBValue::Array(args.to_vec()))
}

fn builtin_createobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CreateObject")?;
    let prog_id = to_arg_string(&args[0]);
    match prog_id.to_uppercase().as_str() {
        "SCRIPTING.DICTIONARY" => Ok(VBValue::Object(Box::new(Dictionary::new()))),
        _ => Err(VBSErrorType::NotImplementedError.into_error(
            format!("CreateObject('{}') is not implemented", prog_id)
        )),
    }
}

fn builtin_len(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Len")?;
    let s = match &args[0] {
        VBValue::String(s) => s.len() as f64,
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::Number(0.0)),
        v => v.to_string().len() as f64,
    };
    Ok(VBValue::Number(s))
}

fn builtin_ucase(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "UCase")?;
    let s = match &args[0] {
        VBValue::String(s) => s.to_uppercase(),
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::String("".to_string())),
        v => v.to_string().to_uppercase(),
    };
    Ok(VBValue::String(s))
}

fn builtin_lcase(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "LCase")?;
    let s = match &args[0] {
        VBValue::String(s) => s.to_lowercase(),
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::String("".to_string())),
        v => v.to_string().to_lowercase(),
    };
    Ok(VBValue::String(s))
}

fn builtin_mid(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "Mid")?;
    let s = to_arg_string(&args[0]);
    let start = to_arg_f64(&args[1]) as usize;
    let length = if args.len() >= 3 {
        Some(to_arg_f64(&args[2]) as usize)
    } else {
        None
    };

    if start < 1 || start > s.len() {
        return Ok(VBValue::String("".to_string()));
    }
    let start_idx = start - 1;
    match length {
        Some(len) => {
            let end = (start_idx + len).min(s.len());
            Ok(VBValue::String(s[start_idx..end].to_string()))
        }
        None => Ok(VBValue::String(s[start_idx..].to_string())),
    }
}

fn builtin_left(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "Left")?;
    let s = to_arg_string(&args[0]);
    let count = to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[..count].to_string()))
}

fn builtin_right(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "Right")?;
    let s = to_arg_string(&args[0]);
    let count = to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[s.len() - count..].to_string()))
}

fn builtin_trim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Trim")?;
    let s = to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim().to_string()))
}

fn builtin_cint(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CInt")?;
    let n = to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.round()))
}

fn builtin_cstr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CStr")?;
    Ok(VBValue::String(to_arg_string(&args[0])))
}

fn builtin_abs(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Abs")?;
    let n = to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.abs()))
}

fn builtin_isnull(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsNull")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Null)))
}

fn builtin_isempty(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsEmpty")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Empty)))
}

fn builtin_instr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    // InStr([start, ]string1, string2)
    let (start, s1, s2) = if args.len() == 3 {
        (to_arg_f64(&args[0]) as usize, to_arg_string(&args[1]), to_arg_string(&args[2]))
    } else if args.len() == 2 {
        (1usize, to_arg_string(&args[0]), to_arg_string(&args[1]))
    } else {
        return Err(VBSErrorType::ValueError.into_error(
            format!("InStr requires 2 or 3 arguments, got {}", args.len())
        ));
    };

    if start < 1 || start > s1.len() {
        return Ok(VBValue::Number(0.0));
    }
    let search_from = start - 1;
    match s1[search_from..].find(&s2) {
        Some(pos) => Ok(VBValue::Number((search_from + pos + 1) as f64)),
        None => Ok(VBValue::Number(0.0)),
    }
}

fn to_arg_string(val: &VBValue) -> String {
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

fn to_arg_f64(val: &VBValue) -> f64 {
    match val {
        VBValue::Number(n) => *n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty | VBValue::Array(_) | VBValue::Object(_) => 0.0,
    }
}
