use super::fso::FileSystemObject;
use super::value::VBValue;
use super::value_utils;
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
        n if n.eq_ignore_ascii_case("SPLIT") => builtin_split(&args),
        n if n.eq_ignore_ascii_case("JOIN") => builtin_join(&args),
        n if n.eq_ignore_ascii_case("REPLACE") => builtin_replace(&args),
        n if n.eq_ignore_ascii_case("ASC") => builtin_asc(&args),
        n if n.eq_ignore_ascii_case("CHR") => builtin_chr(&args),
        n if n.eq_ignore_ascii_case("LTRIM") => builtin_ltrim(&args),
        n if n.eq_ignore_ascii_case("RTRIM") => builtin_rtrim(&args),
        n if n.eq_ignore_ascii_case("SPACE") => builtin_space(&args),
        n if n.eq_ignore_ascii_case("STRING") => builtin_string(&args),
        n if n.eq_ignore_ascii_case("STRREVERSE") => builtin_strreverse(&args),
        n if n.eq_ignore_ascii_case("INSTRREV") => builtin_instrrev(&args),
        n if n.eq_ignore_ascii_case("ISNUMERIC") => builtin_isnumeric(&args),
        n if n.eq_ignore_ascii_case("ISARRAY") => builtin_isarray(&args),
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
    Ok(VBValue::Array(std::sync::Arc::new(args.to_vec())))
}

fn builtin_createobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CreateObject")?;
    let prog_id = value_utils::to_arg_string(&args[0]);
    match prog_id.to_uppercase().as_str() {
        "SCRIPTING.DICTIONARY" => Ok(VBValue::Object(Box::new(Dictionary::new()))),
        "SCRIPTING.FILESYSTEMOBJECT" => Ok(VBValue::Object(Box::new(FileSystemObject::new()))),
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
    let s = value_utils::to_arg_string(&args[0]);
    let start = value_utils::to_arg_f64(&args[1]) as usize;
    let length = if args.len() >= 3 {
        Some(value_utils::to_arg_f64(&args[2]) as usize)
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
    let s = value_utils::to_arg_string(&args[0]);
    let count = value_utils::to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[..count].to_string()))
}

fn builtin_right(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "Right")?;
    let s = value_utils::to_arg_string(&args[0]);
    let count = value_utils::to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[s.len() - count..].to_string()))
}

fn builtin_trim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Trim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim().to_string()))
}

fn builtin_cint(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CInt")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.round()))
}

fn builtin_cstr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CStr")?;
    Ok(VBValue::String(value_utils::to_arg_string(&args[0])))
}

fn builtin_abs(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Abs")?;
    let n = value_utils::to_arg_f64(&args[0]);
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
        (value_utils::to_arg_f64(&args[0]) as usize, value_utils::to_arg_string(&args[1]), value_utils::to_arg_string(&args[2]))
    } else if args.len() == 2 {
        (1usize, value_utils::to_arg_string(&args[0]), value_utils::to_arg_string(&args[1]))
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

fn builtin_split(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "Split")?;
    let s = value_utils::to_arg_string(&args[0]);
    let delimiter = if args.len() >= 2 {
        value_utils::to_arg_string(&args[1])
    } else {
        " ".to_string()
    };
    let count = if args.len() >= 3 {
        value_utils::to_arg_f64(&args[2]) as i64
    } else {
        -1
    };

    let parts: Vec<VBValue> = if delimiter.is_empty() {
        vec![VBValue::String(s)]
    } else {
        let split_result: Vec<&str> = if count < 0 {
            s.split(&delimiter).collect()
        } else {
            let mut result = Vec::new();
            let mut remaining = s.as_str();
            for _ in 0..(count - 1).max(0) {
                match remaining.find(&delimiter) {
                    Some(pos) => {
                        result.push(&remaining[..pos]);
                        remaining = &remaining[pos + delimiter.len()..];
                    }
                    None => break,
                }
            }
            if count > 0 {
                result.push(remaining);
            }
            result
        };
        split_result.into_iter().map(|p| VBValue::String(p.to_string())).collect()
    };
    Ok(VBValue::Array(std::sync::Arc::new(parts)))
}

fn builtin_join(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "Join")?;
    let delimiter = if args.len() >= 2 {
        value_utils::to_arg_string(&args[1])
    } else {
        " ".to_string()
    };
    let arr = match &args[0] {
        VBValue::Array(a) => a,
        _ => return Err(VBSErrorType::ValueError.into_error(
            "Join requires an array as first argument".to_string()
        )),
    };
    let strings: Vec<String> = arr.iter().map(|v| value_utils::to_arg_string(v)).collect();
    Ok(VBValue::String(strings.join(&delimiter)))
}

fn builtin_replace(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 3, "Replace")?;
    let s = value_utils::to_arg_string(&args[0]);
    let find = value_utils::to_arg_string(&args[1]);
    let replace = value_utils::to_arg_string(&args[2]);
    let start = if args.len() >= 4 {
        value_utils::to_arg_f64(&args[3]) as usize
    } else {
        1
    };
    let count = if args.len() >= 5 {
        value_utils::to_arg_f64(&args[4]) as i64
    } else {
        -1
    };

    if find.is_empty() {
        return Ok(VBValue::String(s));
    }
    if start < 1 || start > s.len() {
        return Ok(VBValue::String("".to_string()));
    }
    let search_from = start - 1;
    let search_in = &s[search_from..];

    let mut result = String::new();
    let mut remaining = search_in;
    let mut replacements = 0i64;

    loop {
        if count >= 0 && replacements >= count {
            result.push_str(remaining);
            break;
        }
        match remaining.find(&find) {
            Some(pos) => {
                result.push_str(&remaining[..pos]);
                result.push_str(&replace);
                remaining = &remaining[pos + find.len()..];
                replacements += 1;
            }
            None => {
                result.push_str(remaining);
                break;
            }
        }
    }
    Ok(VBValue::String(result))
}

fn builtin_asc(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Asc")?;
    let s = value_utils::to_arg_string(&args[0]);
    if s.is_empty() {
        return Err(VBSErrorType::ValueError.into_error(
            "Asc requires a non-empty string".to_string()
        ));
    }
    let code = s.chars().next().unwrap() as u32;
    Ok(VBValue::Number(code as f64))
}

fn builtin_chr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Chr")?;
    let code = value_utils::to_arg_f64(&args[0]) as u32;
    match char::from_u32(code) {
        Some(c) => Ok(VBValue::String(c.to_string())),
        None => Err(VBSErrorType::ValueError.into_error(
            format!("Invalid character code: {}", code)
        )),
    }
}

fn builtin_ltrim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "LTrim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim_start().to_string()))
}

fn builtin_rtrim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "RTrim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim_end().to_string()))
}

fn builtin_space(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Space")?;
    let count = value_utils::to_arg_f64(&args[0]) as usize;
    Ok(VBValue::String(" ".repeat(count)))
}

fn builtin_string(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "String")?;
    let count = value_utils::to_arg_f64(&args[0]) as usize;
    let ch = match &args[1] {
        VBValue::Number(n) => {
            let code = *n as u32;
            char::from_u32(code).unwrap_or(' ')
        }
        VBValue::String(s) => s.chars().next().unwrap_or(' '),
        _ => ' ',
    };
    Ok(VBValue::String(ch.to_string().repeat(count)))
}

fn builtin_strreverse(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "StrReverse")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.chars().rev().collect()))
}

fn builtin_instrrev(args: &[VBValue]) -> Result<VBValue, VBSError> {
    // InStrRev(string1, string2[, start[, compare]])
    expect_min_args(args, 2, "InStrRev")?;
    let s1 = value_utils::to_arg_string(&args[0]);
    let s2 = value_utils::to_arg_string(&args[1]);
    let start = if args.len() >= 3 {
        value_utils::to_arg_f64(&args[2]) as usize
    } else {
        s1.len()
    };

    if s2.is_empty() {
        return Ok(VBValue::Number(start as f64));
    }
    let end = start.min(s1.len());
    let search_in = &s1[..end];
    match search_in.rfind(&s2) {
        Some(pos) => Ok(VBValue::Number((pos + 1) as f64)),
        None => Ok(VBValue::Number(0.0)),
    }
}

fn builtin_isnumeric(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsNumeric")?;
    let result = match &args[0] {
        VBValue::Number(_) => true,
        VBValue::String(s) => s.parse::<f64>().is_ok() && !s.is_empty(),
        VBValue::Boolean(_) => false,
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(_) | VBValue::Object(_) => false,
    };
    Ok(VBValue::Boolean(result))
}

fn builtin_isarray(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsArray")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Array(_))))
}


