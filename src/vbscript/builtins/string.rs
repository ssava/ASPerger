use super::expect_arg_count;
use super::expect_min_args;
use crate::vbscript::value::VBValue;
use crate::vbscript::value_utils;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};

pub(super) fn builtin_len(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Len")?;
    let s = match &args[0] {
        VBValue::String(s) => s.len() as f64,
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::Number(0.0)),
        v => v.to_string().len() as f64,
    };
    Ok(VBValue::Number(s))
}

pub(super) fn builtin_ucase(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "UCase")?;
    let s = match &args[0] {
        VBValue::String(s) => s.to_uppercase(),
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::String("".into())),
        v => v.to_string().to_uppercase(),
    };
    Ok(VBValue::String(s.into()))
}

pub(super) fn builtin_lcase(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "LCase")?;
    let s = match &args[0] {
        VBValue::String(s) => s.to_lowercase(),
        VBValue::Null => return Ok(VBValue::Null),
        VBValue::Empty => return Ok(VBValue::String("".into())),
        v => v.to_string().to_lowercase(),
    };
    Ok(VBValue::String(s.into()))
}

pub(super) fn builtin_mid(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "Mid")?;
    let s = value_utils::to_arg_string(&args[0]);
    let start = value_utils::to_arg_f64(&args[1]) as usize;
    let length = if args.len() >= 3 {
        Some(value_utils::to_arg_f64(&args[2]) as usize)
    } else {
        None
    };

    if start < 1 || start > s.len() {
        return Ok(VBValue::String("".into()));
    }
    let start_idx = start - 1;
    match length {
        Some(len) => {
            let end = (start_idx + len).min(s.len());
            Ok(VBValue::String(s[start_idx..end].to_string().into()))
        }
        None => Ok(VBValue::String(s[start_idx..].to_string().into())),
    }
}

pub(super) fn builtin_left(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "Left")?;
    let s = value_utils::to_arg_string(&args[0]);
    let count = value_utils::to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[..count].to_string().into()))
}

pub(super) fn builtin_right(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "Right")?;
    let s = value_utils::to_arg_string(&args[0]);
    let count = value_utils::to_arg_f64(&args[1]) as usize;
    let count = count.min(s.len());
    Ok(VBValue::String(s[s.len() - count..].to_string().into()))
}

pub(super) fn builtin_trim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Trim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim().to_string().into()))
}

pub(super) fn builtin_instr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    let (start, s1, s2) = if args.len() == 3 {
        (
            value_utils::to_arg_f64(&args[0]) as usize,
            value_utils::to_arg_string(&args[1]),
            value_utils::to_arg_string(&args[2]),
        )
    } else if args.len() == 2 {
        (
            1usize,
            value_utils::to_arg_string(&args[0]),
            value_utils::to_arg_string(&args[1]),
        )
    } else {
        return Err(VBSErrorType::ValueError.into_error(format!(
            "InStr requires 2 or 3 arguments, got {}",
            args.len()
        )));
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

pub(super) fn builtin_split(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
        vec![VBValue::String(s.into())]
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
        split_result
            .into_iter()
            .map(|p| VBValue::String(p.to_string().into()))
            .collect()
    };
    Ok(VBValue::Array(std::sync::Arc::new(parts), vec![]))
}

pub(super) fn builtin_join(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "Join")?;
    let delimiter = if args.len() >= 2 {
        value_utils::to_arg_string(&args[1])
    } else {
        " ".to_string()
    };
    let arr = match &args[0] {
        VBValue::Array(a, _) => a,
        _ => {
            return Err(VBSErrorType::ValueError
                .into_error("Join requires an array as first argument".to_string()))
        }
    };
    let strings: Vec<String> = arr.iter().map(value_utils::to_arg_string).collect();
    Ok(VBValue::String(strings.join(&delimiter).into()))
}

pub(super) fn builtin_replace(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 3, "Replace")?;
    let s = value_utils::to_arg_string(&args[0]);
    let find = value_utils::to_arg_string(&args[1]);
    let replace_with = value_utils::to_arg_string(&args[2]);
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
        return Ok(VBValue::String(s.into()));
    }
    if start < 1 || start > s.len() {
        return Ok(VBValue::String("".into()));
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
                result.push_str(&replace_with);
                remaining = &remaining[pos + find.len()..];
                replacements += 1;
            }
            None => {
                result.push_str(remaining);
                break;
            }
        }
    }
    Ok(VBValue::String(result.into()))
}

pub(super) fn builtin_asc(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Asc")?;
    let s = value_utils::to_arg_string(&args[0]);
    if s.is_empty() {
        return Err(
            VBSErrorType::ValueError.into_error("Asc requires a non-empty string".to_string())
        );
    }
    let code = s.chars().next().unwrap() as u32;
    Ok(VBValue::Number(code as f64))
}

pub(super) fn builtin_chr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Chr")?;
    let code = value_utils::to_arg_f64(&args[0]) as u32;
    match char::from_u32(code) {
        Some(c) => Ok(VBValue::String(c.to_string().into())),
        None => {
            Err(VBSErrorType::ValueError.into_error(format!("Invalid character code: {}", code)))
        }
    }
}

pub(super) fn builtin_ltrim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "LTrim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim_start().to_string().into()))
}

pub(super) fn builtin_rtrim(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "RTrim")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.trim_end().to_string().into()))
}

pub(super) fn builtin_space(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Space")?;
    let count = value_utils::to_arg_f64(&args[0]) as usize;
    Ok(VBValue::String(" ".repeat(count).into()))
}

pub(super) fn builtin_string(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
    Ok(VBValue::String(ch.to_string().repeat(count).into()))
}

pub(super) fn builtin_strreverse(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "StrReverse")?;
    let s = value_utils::to_arg_string(&args[0]);
    Ok(VBValue::String(s.chars().rev().collect::<String>().into()))
}

pub(super) fn builtin_instrrev(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_strcomp(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "StrComp")?;
    let s1 = value_utils::to_arg_string(&args[0]);
    let s2 = value_utils::to_arg_string(&args[1]);
    let compare = if args.len() >= 3 {
        value_utils::to_arg_f64(&args[2]) as i32
    } else {
        0
    };
    let result = if compare == 1 {
        s1.to_lowercase().cmp(&s2.to_lowercase()) as i32
    } else {
        s1.cmp(&s2) as i32
    };
    Ok(VBValue::Number(result as f64))
}

fn format_number_internal(num: f64, numdigits: usize) -> String {
    let multiplier = 10_f64.powi(numdigits as i32);
    let sign = if num < 0.0 { "-" } else { "" };
    let abs_num = num.abs();
    let rounded = (abs_num * multiplier).round() / multiplier;
    let int_part = rounded.trunc() as i64;
    let int_str = int_part.to_string();
    let mut with_commas = String::new();
    let chars: Vec<char> = int_str.chars().collect();
    for (count, &c) in chars.iter().rev().enumerate() {
        if count > 0 && count % 3 == 0 {
            with_commas.push(',');
        }
        with_commas.push(c);
    }
    let comma_int: String = with_commas.chars().rev().collect();
    if numdigits > 0 {
        let frac = ((rounded - rounded.trunc()) * multiplier).round() as u64;
        format!("{}{}.{:0width$}", sign, comma_int, frac, width = numdigits)
    } else {
        format!("{}{}", sign, comma_int)
    }
}

pub(super) fn builtin_formatnumber(args: &[VBValue]) -> Result<VBValue, VBSError> {
    if args.is_empty() {
        return Err(VBSErrorType::ValueError
            .into_error("FormatNumber requires at least 1 argument".to_string()));
    }
    let num = value_utils::to_arg_f64(&args[0]);
    let numdigits = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as usize
    } else {
        2
    };
    let _includeleading = if args.len() >= 3 {
        value_utils::to_boolean(&args[2])
    } else {
        true
    };
    let _useparens = if args.len() >= 4 {
        value_utils::to_boolean(&args[3])
    } else {
        false
    };
    let _groupdigits = if args.len() >= 5 {
        value_utils::to_boolean(&args[4])
    } else {
        true
    };
    let formatted = format_number_internal(num, numdigits);
    Ok(VBValue::String(formatted.into()))
}

pub(super) fn builtin_formatcurrency(args: &[VBValue]) -> Result<VBValue, VBSError> {
    if args.is_empty() {
        return Err(VBSErrorType::ValueError
            .into_error("FormatCurrency requires at least 1 argument".to_string()));
    }
    let num = value_utils::to_arg_f64(&args[0]);
    let numdigits = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as usize
    } else {
        2
    };
    let _includeleading = if args.len() >= 3 {
        value_utils::to_boolean(&args[2])
    } else {
        true
    };
    let _useparens = if args.len() >= 4 {
        value_utils::to_boolean(&args[3])
    } else {
        false
    };
    let _groupdigits = if args.len() >= 5 {
        value_utils::to_boolean(&args[4])
    } else {
        true
    };
    let num_str = format_number_internal(num, numdigits);
    let neg = num < 0.0;
    let result = if neg {
        format!("(${})", &num_str[1..])
    } else {
        format!("${}", num_str)
    };
    Ok(VBValue::String(result.into()))
}

pub(super) fn builtin_formatpercent(args: &[VBValue]) -> Result<VBValue, VBSError> {
    if args.is_empty() {
        return Err(VBSErrorType::ValueError
            .into_error("FormatPercent requires at least 1 argument".to_string()));
    }
    let num = value_utils::to_arg_f64(&args[0]) * 100.0;
    let numdigits = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as usize
    } else {
        2
    };
    let _includeleading = if args.len() >= 3 {
        value_utils::to_boolean(&args[2])
    } else {
        true
    };
    let _useparens = if args.len() >= 4 {
        value_utils::to_boolean(&args[3])
    } else {
        false
    };
    let _groupdigits = if args.len() >= 5 {
        value_utils::to_boolean(&args[4])
    } else {
        true
    };
    let num_str = format_number_internal(num, numdigits);
    Ok(VBValue::String(format!("{}%", num_str).into()))
}

pub(super) fn builtin_lset(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "LSet")?;
    let s = value_utils::to_arg_string(&args[0]);
    let length = value_utils::to_arg_f64(&args[1]) as usize;
    if s.len() >= length {
        Ok(VBValue::String(s[..length].to_string().into()))
    } else {
        let mut result = s.clone();
        result.push_str(&" ".repeat(length - s.len()));
        Ok(VBValue::String(result.into()))
    }
}

pub(super) fn builtin_rset(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "RSet")?;
    let s = value_utils::to_arg_string(&args[0]);
    let length = value_utils::to_arg_f64(&args[1]) as usize;
    if s.len() >= length {
        Ok(VBValue::String(s[..length].to_string().into()))
    } else {
        let mut result = " ".repeat(length - s.len());
        result.push_str(&s);
        Ok(VBValue::String(result.into()))
    }
}
