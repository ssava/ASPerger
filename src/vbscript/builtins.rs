//! Built-in VBScript functions: string, math, date/time, type conversion,
//! array, and miscellaneous operations dispatched by name.

use super::fso::FileSystemObject;
use super::value::VBValue;
use super::value_utils;
use super::vbobject::Dictionary;
use super::vbs_error::{VBSError, VBSErrorType};
use chrono::{Datelike, Timelike};
use std::cell::RefCell;

thread_local! {
    static RNG_STATE: RefCell<u32> = const { RefCell::new(0u32) };
    static RNG_LAST: RefCell<f64> = const { RefCell::new(0.0f64) };
}

/// Dispatch a built-in VBScript function call by name.
///
/// Matches case-insensitively against ~80 built-in functions covering
/// string manipulation (`Len`, `Mid`, `Replace`), math (`Abs`, `Sqr`, `Rnd`),
/// date/time (`Now`, `DateAdd`, `DateDiff`), type conversion (`CInt`, `CStr`),
/// array operations (`UBound`, `LBound`, `Split`, `Join`),
/// and I/O (`CreateObject`, `FetchURL`).
pub fn call_builtin(name: &str, args: Vec<VBValue>) -> Result<VBValue, VBSError> {
    match name {
        n if n.eq_ignore_ascii_case("ARRAY") => builtin_array(&args),
        n if n.eq_ignore_ascii_case("FETCHURL") => builtin_fetchurl(&args),
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
        // Date/Time
        n if n.eq_ignore_ascii_case("NOW") => builtin_now(&args),
        n if n.eq_ignore_ascii_case("DATE") => builtin_date(&args),
        n if n.eq_ignore_ascii_case("TIME") => builtin_time(&args),
        n if n.eq_ignore_ascii_case("YEAR") => builtin_year(&args),
        n if n.eq_ignore_ascii_case("MONTH") => builtin_month(&args),
        n if n.eq_ignore_ascii_case("DAY") => builtin_day(&args),
        n if n.eq_ignore_ascii_case("HOUR") => builtin_hour(&args),
        n if n.eq_ignore_ascii_case("MINUTE") => builtin_minute(&args),
        n if n.eq_ignore_ascii_case("SECOND") => builtin_second(&args),
        n if n.eq_ignore_ascii_case("WEEKDAY") => builtin_weekday(&args),
        n if n.eq_ignore_ascii_case("WEEKDAYNAME") => builtin_weekdayname(&args),
        n if n.eq_ignore_ascii_case("MONTHNAME") => builtin_monthname(&args),
        n if n.eq_ignore_ascii_case("DATEADD") => builtin_dateadd(&args),
        n if n.eq_ignore_ascii_case("DATEDIFF") => builtin_datediff(&args),
        n if n.eq_ignore_ascii_case("DATEPART") => builtin_datepart(&args),
        n if n.eq_ignore_ascii_case("DATESERIAL") => builtin_dateserial(&args),
        n if n.eq_ignore_ascii_case("DATEVALUE") => builtin_datevalue(&args),
        n if n.eq_ignore_ascii_case("TIMESERIAL") => builtin_timeserial(&args),
        n if n.eq_ignore_ascii_case("TIMEVALUE") => builtin_timevalue(&args),
        n if n.eq_ignore_ascii_case("TIMER") => builtin_timer(&args),
        n if n.eq_ignore_ascii_case("FORMATDATETIME") => builtin_formatdatetime(&args),
        // Math
        n if n.eq_ignore_ascii_case("RND") => builtin_rnd(&args),
        n if n.eq_ignore_ascii_case("RANDOMIZE") => builtin_randomize(&args),
        n if n.eq_ignore_ascii_case("INT") => builtin_int(&args),
        n if n.eq_ignore_ascii_case("FIX") => builtin_fix(&args),
        n if n.eq_ignore_ascii_case("ROUND") => builtin_round(&args),
        n if n.eq_ignore_ascii_case("SGN") => builtin_sgn(&args),
        n if n.eq_ignore_ascii_case("SQR") => builtin_sqr(&args),
        // Array
        n if n.eq_ignore_ascii_case("UBOUND") => builtin_ubound(&args),
        n if n.eq_ignore_ascii_case("LBOUND") => builtin_lbound(&args),
        n if n.eq_ignore_ascii_case("FILTER") => builtin_filter(&args),
        // Type Conversion
        n if n.eq_ignore_ascii_case("CBOOL") => builtin_cbool(&args),
        n if n.eq_ignore_ascii_case("CBYTE") => builtin_cbyte(&args),
        n if n.eq_ignore_ascii_case("CDATE") => builtin_cdate(&args),
        n if n.eq_ignore_ascii_case("CDBL") => builtin_cdbl(&args),
        n if n.eq_ignore_ascii_case("CLNG") => builtin_clng(&args),
        n if n.eq_ignore_ascii_case("CSNG") => builtin_csng(&args),
        n if n.eq_ignore_ascii_case("CCUR") => builtin_ccur(&args),
        // Other
        n if n.eq_ignore_ascii_case("HEX") => builtin_hex(&args),
        n if n.eq_ignore_ascii_case("OCT") => builtin_oct(&args),
        n if n.eq_ignore_ascii_case("ISDATE") => builtin_isdate(&args),
        n if n.eq_ignore_ascii_case("ISOBJECT") => builtin_isobject(&args),
        n if n.eq_ignore_ascii_case("TYPENAME") => builtin_typename(&args),
        n if n.eq_ignore_ascii_case("VARTYPE") => builtin_vartype(&args),
        // Remaining string functions
        n if n.eq_ignore_ascii_case("STRCOMP") => builtin_strcomp(&args),
        n if n.eq_ignore_ascii_case("FORMATCURRENCY") => builtin_formatcurrency(&args),
        n if n.eq_ignore_ascii_case("FORMATNUMBER") => builtin_formatnumber(&args),
        n if n.eq_ignore_ascii_case("FORMATPERCENT") => builtin_formatpercent(&args),
        n if n.eq_ignore_ascii_case("LSET") => builtin_lset(&args),
        n if n.eq_ignore_ascii_case("RSET") => builtin_rset(&args),
        _ => Err(VBSErrorType::NotImplementedError
            .into_error(format!("Function '{}' is not implemented", name))),
    }
}

fn expect_arg_count(args: &[VBValue], expected: usize, name: &str) -> Result<(), VBSError> {
    if args.len() != expected {
        return Err(VBSErrorType::ValueError.into_error(format!(
            "{} requires {} argument(s), got {}",
            name,
            expected,
            args.len()
        )));
    }
    Ok(())
}

fn expect_min_args(args: &[VBValue], min: usize, name: &str) -> Result<(), VBSError> {
    if args.len() < min {
        return Err(VBSErrorType::ValueError.into_error(format!(
            "{} requires at least {} argument(s), got {}",
            name,
            min,
            args.len()
        )));
    }
    Ok(())
}

fn builtin_array(args: &[VBValue]) -> Result<VBValue, VBSError> {
    Ok(VBValue::Array(std::sync::Arc::new(args.to_vec()), vec![]))
}

fn builtin_createobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CreateObject")?;
    let prog_id = value_utils::to_arg_string(&args[0]);
    match prog_id.to_uppercase().as_str() {
        "ADODB.CONNECTION" => Ok(VBValue::Object(Box::new(
            crate::vbscript::adodb::Connection::new(),
        ))),
        "SCRIPTING.DICTIONARY" => Ok(VBValue::Object(Box::new(Dictionary::new()))),
        "SCRIPTING.FILESYSTEMOBJECT" => Ok(VBValue::Object(Box::new(FileSystemObject::new()))),
        "VBSCRIPT.REGEXP" => Ok(VBValue::Object(Box::new(
            crate::vbscript::regexp::RegExpObject::new(),
        ))),
        _ => Err(VBSErrorType::NotImplementedError
            .into_error(format!("CreateObject('{}') is not implemented", prog_id))),
    }
}

fn builtin_fetchurl(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "FetchURL")?;
    let url = value_utils::to_arg_string(&args[0]);
    let client = reqwest::Client::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .map_err(|e| VBSErrorType::RuntimeError
            .into_error(format!("FetchURL: failed to create client: {e}")))?;
    let rt_handle = tokio::runtime::Handle::current();
    let result = tokio::task::block_in_place(move || {
        rt_handle.block_on(async {
            match client.get(&url).send().await {
                Ok(resp) => match resp.text().await {
                    Ok(body) => Ok(VBValue::String(body)),
                    Err(e) => Err(VBSErrorType::RuntimeError
                        .into_error(format!("FetchURL: failed to read response body: {e}"))),
                },
                Err(e) => Err(VBSErrorType::RuntimeError
                    .into_error(format!("FetchURL: request failed: {e}"))),
            }
        })
    });
    result
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
        split_result
            .into_iter()
            .map(|p| VBValue::String(p.to_string()))
            .collect()
    };
    Ok(VBValue::Array(std::sync::Arc::new(parts), vec![]))
}

fn builtin_join(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
        return Err(
            VBSErrorType::ValueError.into_error("Asc requires a non-empty string".to_string())
        );
    }
    let code = s.chars().next().unwrap() as u32;
    Ok(VBValue::Number(code as f64))
}

fn builtin_chr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Chr")?;
    let code = value_utils::to_arg_f64(&args[0]) as u32;
    match char::from_u32(code) {
        Some(c) => Ok(VBValue::String(c.to_string())),
        None => {
            Err(VBSErrorType::ValueError.into_error(format!("Invalid character code: {}", code)))
        }
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
        VBValue::Array(..) | VBValue::Object(_) => false,
    };
    Ok(VBValue::Boolean(result))
}

fn builtin_isarray(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsArray")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Array(..))))
}

// ======================== OLE Automation Date Helpers ========================

pub(crate) fn ole_auto_to_datetime(serial: f64) -> Option<chrono::NaiveDateTime> {
    if serial.is_nan() || serial.is_infinite() {
        return None;
    }
    let epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)?;
    let days = serial.trunc() as i64;
    let total_secs = (serial.fract().abs() * 86400.0).round() as u32;
    let total_secs = total_secs.min(86399);
    let date = epoch.checked_add_signed(chrono::Duration::days(days))?;
    let time = chrono::NaiveTime::from_num_seconds_from_midnight_opt(total_secs, 0)?;
    Some(chrono::NaiveDateTime::new(date, time))
}

pub(crate) fn datetime_to_ole_auto(dt: chrono::NaiveDateTime) -> f64 {
    let epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let epoch_dt = chrono::NaiveDateTime::new(
        epoch,
        chrono::NaiveTime::from_num_seconds_from_midnight_opt(0, 0).unwrap(),
    );
    let days = dt.signed_duration_since(epoch_dt).num_days() as f64;
    let seconds = dt.num_seconds_from_midnight() as f64;
    days + seconds / 86400.0
}

fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0) {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

fn add_months(dt: chrono::NaiveDateTime, months: i32) -> chrono::NaiveDateTime {
    if months == 0 {
        return dt;
    }
    let total_months = dt.year() * 12 + dt.month() as i32 - 1 + months;
    let new_year = total_months.div_euclid(12);
    let new_month = (total_months.rem_euclid(12) + 1) as u32;
    let max_day = days_in_month(new_year, new_month);
    let new_day = dt.day().min(max_day);
    let date = chrono::NaiveDate::from_ymd_opt(new_year, new_month, new_day).expect("valid date");
    chrono::NaiveDateTime::new(date, dt.time())
}

/// Convert a VBValue to NaiveDateTime: try numeric OLE date first, then string parse.
pub(crate) fn value_to_datetime(val: &VBValue) -> Option<chrono::NaiveDateTime> {
    match val {
        VBValue::Number(n) => ole_auto_to_datetime(*n),
        VBValue::String(s) => try_parse_date(s),
        _ => {
            let s = value_utils::to_arg_string(val);
            try_parse_date(&s)
        }
    }
}

pub(crate) fn try_parse_date(s: &str) -> Option<chrono::NaiveDateTime> {
    let s = s.trim();
    let datetime_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S %p",
    ];
    for fmt in &datetime_formats {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, fmt) {
            return Some(dt);
        }
    }
    let date_formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d %B %Y",
        "%d %b %Y",
        "%B %-d, %Y",
    ];
    for fmt in &date_formats {
        if let Ok(d) = chrono::NaiveDate::parse_from_str(s, fmt) {
            return Some(chrono::NaiveDateTime::new(
                d,
                chrono::NaiveTime::from_num_seconds_from_midnight_opt(0, 0).unwrap(),
            ));
        }
    }
    let time_formats = ["%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"];
    for fmt in &time_formats {
        if let Ok(t) = chrono::NaiveTime::parse_from_str(s, fmt) {
            return Some(chrono::NaiveDateTime::new(
                chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap(),
                t,
            ));
        }
    }
    None
}

// ======================== RNG Helpers ========================

fn rnd_next(state: &mut u32) -> f64 {
    *state = state.wrapping_mul(1103515245).wrapping_add(12345) & 0x7fffffff;
    *state as f64 / 2147483648.0
}

// ======================== Date/Time Functions ========================

fn builtin_now(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::Number(datetime_to_ole_auto(now)))
}

fn builtin_date(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::String(now.format("%m/%d/%Y").to_string()))
}

fn builtin_time(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::String(now.format("%H:%M:%S").to_string()))
}

fn builtin_year(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Year")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.year() as f64))
}

fn builtin_month(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Month")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.month() as f64))
}

fn builtin_day(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Day")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.day() as f64))
}

fn builtin_hour(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Hour")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.hour() as f64))
}

fn builtin_minute(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Minute")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.minute() as f64))
}

fn builtin_second(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Second")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.second() as f64))
}

fn builtin_weekday(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "Weekday")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let firstday = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as u32
    } else {
        1
    };
    let vbs_weekday = dt.weekday().num_days_from_sunday() + 1;
    let result = ((vbs_weekday + 7 - firstday) % 7) + 1;
    Ok(VBValue::Number(result as f64))
}

fn builtin_weekdayname(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "WeekdayName")?;
    let weekday = value_utils::to_arg_f64(&args[0]) as u32;
    let abbreviate = if args.len() >= 2 {
        value_utils::to_boolean(&args[1])
    } else {
        false
    };
    let firstday = if args.len() >= 3 {
        value_utils::to_arg_f64(&args[2]) as u32
    } else {
        1
    };
    if !(1..=7).contains(&weekday) {
        return Err(VBSErrorType::ValueError.into_error("Invalid weekday".to_string()));
    }
    let idx = ((weekday - 1) + (firstday - 1)) % 7;
    let names = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let short = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    if abbreviate {
        Ok(VBValue::String(short[idx as usize].to_string()))
    } else {
        Ok(VBValue::String(names[idx as usize].to_string()))
    }
}

fn builtin_monthname(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "MonthName")?;
    let month = value_utils::to_arg_f64(&args[0]) as usize;
    let abbreviate = if args.len() >= 2 {
        value_utils::to_boolean(&args[1])
    } else {
        false
    };
    if !(1..=12).contains(&month) {
        return Err(VBSErrorType::ValueError.into_error("Invalid month".to_string()));
    }
    let names = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let short = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    if abbreviate {
        Ok(VBValue::String(short[month - 1].to_string()))
    } else {
        Ok(VBValue::String(names[month - 1].to_string()))
    }
}

fn builtin_dateadd(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 3, "DateAdd")?;
    let interval = value_utils::to_arg_string(&args[0]).to_lowercase();
    let number = value_utils::to_arg_f64(&args[1]);
    let date = value_utils::to_arg_f64(&args[2]);
    let dt = ole_auto_to_datetime(date)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let result = match interval.as_str() {
        "yyyy" => add_months(dt, (number as i32) * 12),
        "q" => add_months(dt, (number as i32) * 3),
        "m" => add_months(dt, number as i32),
        "y" | "d" | "w" => {
            let days = number as i64;
            dt.checked_add_signed(chrono::Duration::days(days))
                .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?
        }
        "ww" => {
            let days = (number * 7.0) as i64;
            dt.checked_add_signed(chrono::Duration::days(days))
                .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?
        }
        "h" => {
            let hours = number as i64;
            dt.checked_add_signed(chrono::Duration::hours(hours))
                .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?
        }
        "n" => {
            let mins = number as i64;
            dt.checked_add_signed(chrono::Duration::minutes(mins))
                .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?
        }
        "s" => {
            let secs = number as i64;
            dt.checked_add_signed(chrono::Duration::seconds(secs))
                .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?
        }
        _ => {
            return Err(
                VBSErrorType::RuntimeError.into_error(format!("Invalid interval '{}'", interval))
            )
        }
    };
    Ok(VBValue::Number(datetime_to_ole_auto(result)))
}

fn builtin_datediff(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 3, "DateDiff")?;
    let interval = value_utils::to_arg_string(&args[0]).to_lowercase();
    let date1 = value_utils::to_arg_f64(&args[1]);
    let date2 = value_utils::to_arg_f64(&args[2]);
    let dt1 = ole_auto_to_datetime(date1)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let dt2 = ole_auto_to_datetime(date2)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let result = match interval.as_str() {
        "yyyy" => (dt2.year() - dt1.year()) as f64,
        "q" => {
            ((dt2.year() - dt1.year()) * 12 + (dt2.month() as i32 - dt1.month() as i32)) as f64
                / 3.0
        }
        "m" => ((dt2.year() - dt1.year()) * 12 + dt2.month() as i32 - dt1.month() as i32) as f64,
        "y" | "d" => {
            let dur = dt2.signed_duration_since(dt1);
            dur.num_days() as f64
        }
        "w" => {
            let firstday = if args.len() >= 4 {
                value_utils::to_arg_f64(&args[3]) as u32
            } else {
                1
            };
            let target = (firstday + 6) % 7;
            let mut count = 0i64;
            let (start, end, neg) = if dt1 <= dt2 {
                (dt1, dt2, false)
            } else {
                (dt2, dt1, true)
            };
            let mut current = start;
            while current.date() < end.date() {
                if current.weekday().num_days_from_sunday() == target {
                    count += 1;
                }
                current = current
                    .checked_add_signed(chrono::Duration::days(1))
                    .unwrap();
            }
            if neg {
                -count as f64
            } else {
                count as f64
            }
        }
        "ww" => {
            let dur = dt2.signed_duration_since(dt1);
            dur.num_days() as f64 / 7.0
        }
        "h" => {
            let dur = dt2.signed_duration_since(dt1);
            dur.num_hours() as f64
        }
        "n" => {
            let dur = dt2.signed_duration_since(dt1);
            dur.num_minutes() as f64
        }
        "s" => {
            let dur = dt2.signed_duration_since(dt1);
            dur.num_seconds() as f64
        }
        _ => {
            return Err(
                VBSErrorType::RuntimeError.into_error(format!("Invalid interval '{}'", interval))
            )
        }
    };
    Ok(VBValue::Number(result))
}

fn builtin_datepart(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "DatePart")?;
    let interval = value_utils::to_arg_string(&args[0]).to_lowercase();
    let dt = value_to_datetime(&args[1])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let result = match interval.as_str() {
        "yyyy" => dt.year() as f64,
        "q" => ((dt.month() - 1) / 3 + 1) as f64,
        "m" => dt.month() as f64,
        "y" => dt.ordinal() as f64,
        "d" => dt.day() as f64,
        "w" => (dt.weekday().num_days_from_sunday() + 1) as f64,
        "ww" => {
            let _firstday = if args.len() >= 3 {
                value_utils::to_arg_f64(&args[2]) as u32
            } else {
                1
            };
            let jan1 = chrono::NaiveDate::from_ymd_opt(dt.year(), 1, 1).unwrap();
            let jan1_wd = jan1.weekday().num_days_from_sunday() as i32;
            let day_of_year = dt.ordinal() as i32 - 1;
            let week = (day_of_year + jan1_wd) / 7 + 1;
            week as f64
        }
        "h" => dt.hour() as f64,
        "n" => dt.minute() as f64,
        "s" => dt.second() as f64,
        _ => {
            return Err(
                VBSErrorType::RuntimeError.into_error(format!("Invalid interval '{}'", interval))
            )
        }
    };
    Ok(VBValue::Number(result))
}

fn builtin_dateserial(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 3, "DateSerial")?;
    let year = value_utils::to_arg_f64(&args[0]) as i32;
    let month = value_utils::to_arg_f64(&args[1]) as i32;
    let day = value_utils::to_arg_f64(&args[2]) as i32;
    let m = month - 1;
    let y = year as i64 + m.div_euclid(12) as i64;
    let mo = m.rem_euclid(12) + 1;
    let date = chrono::NaiveDate::from_ymd_opt(y as i32, mo as u32, 1)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let date = date
        .checked_add_signed(chrono::Duration::days((day - 1) as i64))
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Date overflow".to_string()))?;
    let dt = chrono::NaiveDateTime::new(
        date,
        chrono::NaiveTime::from_num_seconds_from_midnight_opt(0, 0).unwrap(),
    );
    Ok(VBValue::Number(datetime_to_ole_auto(dt)))
}

fn builtin_datevalue(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "DateValue")?;
    let s = value_utils::to_arg_string(&args[0]);
    let dt = try_parse_date(&s)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date string".to_string()))?;
    let date = chrono::NaiveDate::from_ymd_opt(dt.year(), dt.month(), dt.day()).unwrap();
    let result = datetime_to_ole_auto(chrono::NaiveDateTime::new(
        date,
        chrono::NaiveTime::from_num_seconds_from_midnight_opt(0, 0).unwrap(),
    ));
    Ok(VBValue::Number(result))
}

fn builtin_timeserial(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 3, "TimeSerial")?;
    let hour = value_utils::to_arg_f64(&args[0]) as i64;
    let minute = value_utils::to_arg_f64(&args[1]) as i64;
    let second = value_utils::to_arg_f64(&args[2]) as i64;
    let total_seconds = hour * 3600 + minute * 60 + second;
    let s = ((total_seconds % 86400) + 86400) % 86400;
    let time = chrono::NaiveTime::from_num_seconds_from_midnight_opt(s as u32, 0).unwrap();
    let dt =
        chrono::NaiveDateTime::new(chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap(), time);
    Ok(VBValue::Number(datetime_to_ole_auto(dt)))
}

fn builtin_timevalue(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "TimeValue")?;
    let s = value_utils::to_arg_string(&args[0]);
    let time = chrono::NaiveTime::parse_from_str(&s, "%H:%M:%S")
        .or_else(|_| chrono::NaiveTime::parse_from_str(&s, "%H:%M"))
        .or_else(|_| chrono::NaiveTime::parse_from_str(&s, "%I:%M:%S %p"))
        .or_else(|_| chrono::NaiveTime::parse_from_str(&s, "%I:%M %p"))
        .or_else(|_| {
            try_parse_date(&s)
                .map(|dt| {
                    chrono::NaiveTime::from_num_seconds_from_midnight_opt(
                        dt.num_seconds_from_midnight(),
                        0,
                    )
                    .unwrap()
                })
                .ok_or(std::io::Error::other(""))
        })
        .map_err(|_| VBSErrorType::RuntimeError.into_error("Invalid time string".to_string()))?;
    let dt =
        chrono::NaiveDateTime::new(chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap(), time);
    Ok(VBValue::Number(datetime_to_ole_auto(dt)))
}

fn builtin_timer(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    let seconds = now.num_seconds_from_midnight() as f64;
    let nanos = now.and_utc().timestamp_subsec_nanos() as f64;
    Ok(VBValue::Number(seconds + nanos / 1_000_000_000.0))
}

fn builtin_formatdatetime(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "FormatDateTime")?;
    let n = value_utils::to_arg_f64(&args[0]);
    let namedformat = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as i32
    } else {
        0
    };
    let dt = ole_auto_to_datetime(n)
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    let s = match namedformat {
        0 => format!("{} {}", dt.format("%m/%d/%Y"), dt.format("%H:%M:%S")),
        1 => dt.format("%A, %B %-d, %Y").to_string(),
        2 => dt.format("%m/%d/%Y").to_string(),
        3 => {
            let h12 = dt.hour12().1;
            let ampm = if dt.hour() < 12 { "AM" } else { "PM" };
            format!("{:02}:{:02}:{:02} {}", h12, dt.minute(), dt.second(), ampm)
        }
        4 => dt.format("%H:%M").to_string(),
        _ => format!("{} {}", dt.format("%m/%d/%Y"), dt.format("%H:%M:%S")),
    };
    Ok(VBValue::String(s))
}

// ======================== Math Functions ========================

fn builtin_rnd(args: &[VBValue]) -> Result<VBValue, VBSError> {
    let seed = if args.is_empty() {
        None
    } else {
        Some(value_utils::to_arg_f64(&args[0]))
    };
    RNG_STATE.with(|state| {
        RNG_LAST.with(|last| {
            let mut s = state.borrow_mut();
            let mut l = last.borrow_mut();
            match seed {
                Some(n) if n < 0.0 => {
                    *s = ((-n) as u32) & 0x7fffffff;
                    let val = rnd_next(&mut s);
                    *l = val;
                    Ok(VBValue::Number(val))
                }
                Some(0.0) => Ok(VBValue::Number(*l)),
                _ => {
                    let val = rnd_next(&mut s);
                    *l = val;
                    Ok(VBValue::Number(val))
                }
            }
        })
    })
}

fn builtin_randomize(args: &[VBValue]) -> Result<VBValue, VBSError> {
    let seed = if args.is_empty() {
        let now = chrono::Local::now().naive_local();
        now.and_utc().timestamp_subsec_nanos()
    } else {
        value_utils::to_arg_f64(&args[0]) as u32
    };
    RNG_STATE.with(|state| {
        *state.borrow_mut() = seed & 0x7fffffff;
    });
    Ok(VBValue::Null)
}

fn builtin_int(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Int")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.floor()))
}

fn builtin_fix(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Fix")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.trunc()))
}

fn builtin_round(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "Round")?;
    let n = value_utils::to_arg_f64(&args[0]);
    let places = if args.len() >= 2 {
        value_utils::to_arg_f64(&args[1]) as i32
    } else {
        0
    };
    let multiplier = 10_f64.powi(places);
    Ok(VBValue::Number((n * multiplier).round() / multiplier))
}

fn builtin_sgn(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Sgn")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n > 0.0 {
        Ok(VBValue::Number(1.0))
    } else if n < 0.0 {
        Ok(VBValue::Number(-1.0))
    } else {
        Ok(VBValue::Number(0.0))
    }
}

fn builtin_sqr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Sqr")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n < 0.0 {
        return Err(VBSErrorType::RuntimeError
            .into_error("Cannot calculate square root of a negative number".to_string()));
    }
    Ok(VBValue::Number(n.sqrt()))
}

// ======================== Array Functions ========================

fn builtin_ubound(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "UBound")?;
    let arr = match &args[0] {
        VBValue::Array(a, _) => a,
        _ => {
            return Err(VBSErrorType::ValueError.into_error("UBound requires an array".to_string()))
        }
    };
    if arr.is_empty() {
        return Err(VBSErrorType::RuntimeError.into_error("Subscript out of range".to_string()));
    }
    if args.len() >= 2 {
        let _dim = value_utils::to_arg_f64(&args[1]) as usize;
    }
    Ok(VBValue::Number((arr.len() - 1) as f64))
}

fn builtin_lbound(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "LBound")?;
    match &args[0] {
        VBValue::Array(a, _) => {
            if a.is_empty() {
                return Err(
                    VBSErrorType::RuntimeError.into_error("Subscript out of range".to_string())
                );
            }
        }
        _ => {
            return Err(VBSErrorType::ValueError.into_error("LBound requires an array".to_string()))
        }
    };
    Ok(VBValue::Number(0.0))
}

fn builtin_filter(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "Filter")?;
    let arr = match &args[0] {
        VBValue::Array(a, _) => a,
        _ => {
            return Err(VBSErrorType::ValueError.into_error("Filter requires an array".to_string()))
        }
    };
    let match_str = value_utils::to_arg_string(&args[1]);
    let include = if args.len() >= 3 {
        value_utils::to_boolean(&args[2])
    } else {
        true
    };
    let compare = if args.len() >= 4 {
        value_utils::to_arg_f64(&args[3]) as i32
    } else {
        0
    };
    let result: Vec<VBValue> = arr
        .iter()
        .filter(|v| {
            let s = value_utils::to_arg_string(v);
            let found = if compare == 1 {
                s.to_lowercase().contains(&match_str.to_lowercase())
            } else {
                s.contains(&match_str)
            };
            if include {
                found
            } else {
                !found
            }
        })
        .cloned()
        .collect();
    Ok(VBValue::Array(std::sync::Arc::new(result), vec![]))
}

// ======================== Type Conversion Functions ========================

fn builtin_cbool(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CBool")?;
    Ok(VBValue::Boolean(value_utils::to_boolean(&args[0])))
}

fn builtin_cbyte(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CByte")?;
    let n = value_utils::to_arg_f64(&args[0]).round();
    if !(0.0..=255.0).contains(&n) {
        return Err(VBSErrorType::RuntimeError.into_error("Overflow".to_string()));
    }
    Ok(VBValue::Number(n))
}

fn builtin_cdate(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CDate")?;
    match &args[0] {
        VBValue::Number(n) => Ok(VBValue::Number(*n)),
        VBValue::String(s) => {
            let dt = try_parse_date(s).ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("Invalid date string".to_string())
            })?;
            Ok(VBValue::Number(datetime_to_ole_auto(dt)))
        }
        v => {
            let s = value_utils::to_arg_string(v);
            let dt = try_parse_date(&s).ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("Invalid date string".to_string())
            })?;
            Ok(VBValue::Number(datetime_to_ole_auto(dt)))
        }
    }
}

fn builtin_cdbl(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CDbl")?;
    Ok(VBValue::Number(value_utils::to_arg_f64(&args[0])))
}

fn builtin_clng(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CLng")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.trunc()))
}

fn builtin_csng(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CSng")?;
    Ok(VBValue::Number(value_utils::to_arg_f64(&args[0])))
}

fn builtin_ccur(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CCur")?;
    let n = value_utils::to_arg_f64(&args[0]);
    // Currency is fixed-point with 4 decimal places
    Ok(VBValue::Number((n * 10000.0).round() / 10000.0))
}

// ======================== Other Functions ========================

fn builtin_hex(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Hex")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n == 0.0 {
        return Ok(VBValue::String("0".to_string()));
    }
    let n = n as i64;
    Ok(VBValue::String(format!("{:X}", n)))
}

fn builtin_oct(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Oct")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n == 0.0 {
        return Ok(VBValue::String("0".to_string()));
    }
    let n = n as i64;
    Ok(VBValue::String(format!("{:o}", n)))
}

fn builtin_isdate(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsDate")?;
    match &args[0] {
        VBValue::Number(_) => Ok(VBValue::Boolean(true)),
        VBValue::String(s) => Ok(VBValue::Boolean(try_parse_date(s).is_some())),
        _ => Ok(VBValue::Boolean(false)),
    }
}

fn builtin_isobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsObject")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Object(_))))
}

fn builtin_typename(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "TypeName")?;
    let name = match &args[0] {
        VBValue::String(_) => "String",
        VBValue::Boolean(_) => "Boolean",
        VBValue::Null => "Null",
        VBValue::Empty => "Empty",
        VBValue::Array(..) => "Array",
        VBValue::Object(obj) => obj.type_name(),
        VBValue::Number(n) => {
            if n.fract() == 0.0 {
                let n = *n as i64;
                if (-32768..=32767).contains(&n) {
                    "Integer"
                } else if (-2147483648..=2147483647).contains(&n) {
                    "Long"
                } else {
                    "Double"
                }
            } else {
                "Double"
            }
        }
    };
    Ok(VBValue::String(name.to_string()))
}

fn builtin_vartype(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "VarType")?;
    let vt = match &args[0] {
        VBValue::Empty => 0,
        VBValue::Null => 1,
        VBValue::Number(n) => {
            if n.fract() == 0.0 {
                let n = *n as i64;
                if (-32768..=32767).contains(&n) {
                    2
                } else if (-2147483648..=2147483647).contains(&n) {
                    3
                } else {
                    5
                }
            } else {
                5
            }
        }
        VBValue::String(_) => 8,
        VBValue::Boolean(_) => 11,
        VBValue::Object(_) => 9,
        VBValue::Array(..) => 8204,
    };
    Ok(VBValue::Number(vt as f64))
}

// ======================== Remaining String Functions ========================

fn builtin_strcomp(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

fn builtin_formatnumber(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
    Ok(VBValue::String(formatted))
}

fn builtin_formatcurrency(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
    Ok(VBValue::String(result))
}

fn builtin_formatpercent(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
    Ok(VBValue::String(format!("{}%", num_str)))
}

fn builtin_lset(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "LSet")?;
    let s = value_utils::to_arg_string(&args[0]);
    let length = value_utils::to_arg_f64(&args[1]) as usize;
    if s.len() >= length {
        Ok(VBValue::String(s[..length].to_string()))
    } else {
        let mut result = s.clone();
        result.push_str(&" ".repeat(length - s.len()));
        Ok(VBValue::String(result))
    }
}

fn builtin_rset(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 2, "RSet")?;
    let s = value_utils::to_arg_string(&args[0]);
    let length = value_utils::to_arg_f64(&args[1]) as usize;
    if s.len() >= length {
        Ok(VBValue::String(s[..length].to_string()))
    } else {
        let mut result = " ".repeat(length - s.len());
        result.push_str(&s);
        Ok(VBValue::String(result))
    }
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
