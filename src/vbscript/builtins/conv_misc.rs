use super::expect_arg_count;
use super::expect_min_args;
use crate::vbscript::fso::FileSystemObject;
use crate::vbscript::value::VBValue;
use crate::vbscript::value_utils;
use crate::vbscript::vbobject::Dictionary;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};

pub(super) fn builtin_createobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_fetchurl(args: &[VBValue]) -> Result<VBValue, VBSError> {
    #[cfg(not(feature = "fetchurl"))]
    {
        let _ = args;
        Err(VBSErrorType::NotImplementedError
            .into_error("FetchURL requires the 'fetchurl' feature".to_string()))
    }
    #[cfg(feature = "fetchurl")]
    {
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
                        Ok(body) => Ok(VBValue::String(body.into())),
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
}

pub(super) fn builtin_cint(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CInt")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.round()))
}

pub(super) fn builtin_cstr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CStr")?;
    Ok(VBValue::String(value_utils::to_arg_string(&args[0]).into()))
}

pub(super) fn builtin_isnull(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsNull")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Null)))
}

pub(super) fn builtin_isempty(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsEmpty")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Empty)))
}

pub(super) fn builtin_isnumeric(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_isarray(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsArray")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Array(..))))
}

pub(super) fn builtin_ubound(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "UBound")?;
    let (items, dims) = match &args[0] {
        VBValue::Array(a, d) => (a, d),
        _ => {
            return Err(VBSErrorType::ValueError
                .into_error("UBound requires an array".to_string()))
        }
    };
    if items.is_empty() {
        return Err(
            VBSErrorType::RuntimeError.into_error("Subscript out of range".to_string())
        );
    }
    if args.len() >= 2 {
        let dim = value_utils::to_arg_f64(&args[1]) as usize;
        if dim < 1 || dim > dims.len() {
            return Err(VBSErrorType::RuntimeError
                .into_error("Subscript out of range".to_string()));
        }
        if dims.is_empty() {
            return Ok(VBValue::Number((items.len() - 1) as f64));
        }
        return Ok(VBValue::Number(dims[dim - 1] as f64));
    }
    if dims.is_empty() {
        return Ok(VBValue::Number((items.len() - 1) as f64));
    }
    Ok(VBValue::Number(dims[0] as f64))
}

pub(super) fn builtin_lbound(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 1, "LBound")?;
    let (items, dims) = match &args[0] {
        VBValue::Array(a, d) => (a, d),
        _ => {
            return Err(VBSErrorType::ValueError
                .into_error("LBound requires an array".to_string()))
        }
    };
    if items.is_empty() {
        return Err(
            VBSErrorType::RuntimeError.into_error("Subscript out of range".to_string())
        );
    }
    if args.len() >= 2 {
        let dim = value_utils::to_arg_f64(&args[1]) as usize;
        if dim < 1 || dim > dims.len() {
            return Err(VBSErrorType::RuntimeError
                .into_error("Subscript out of range".to_string()));
        }
    }
    Ok(VBValue::Number(0.0))
}

pub(super) fn builtin_filter(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_min_args(args, 2, "Filter")?;
    let arr = match &args[0] {
        VBValue::Array(a, _) => a,
        _ => {
            return Err(VBSErrorType::ValueError
                .into_error("Filter requires an array".to_string()))
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

pub(super) fn builtin_cbool(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CBool")?;
    Ok(VBValue::Boolean(value_utils::to_boolean(&args[0])))
}

pub(super) fn builtin_cbyte(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CByte")?;
    let n = value_utils::to_arg_f64(&args[0]).round();
    if !(0.0..=255.0).contains(&n) {
        return Err(VBSErrorType::RuntimeError.into_error("Overflow".to_string()));
    }
    Ok(VBValue::Number(n))
}

pub(super) fn builtin_cdate(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CDate")?;
    match &args[0] {
        VBValue::Number(n) => Ok(VBValue::Number(*n)),
        VBValue::String(s) => {
            let dt = super::datetime::try_parse_date(s).ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("Invalid date string".to_string())
            })?;
            Ok(VBValue::Number(super::datetime::datetime_to_ole_auto(dt)))
        }
        v => {
            let s = value_utils::to_arg_string(v);
            let dt = super::datetime::try_parse_date(&s).ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("Invalid date string".to_string())
            })?;
            Ok(VBValue::Number(super::datetime::datetime_to_ole_auto(dt)))
        }
    }
}

pub(super) fn builtin_cdbl(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CDbl")?;
    Ok(VBValue::Number(value_utils::to_arg_f64(&args[0])))
}

pub(super) fn builtin_clng(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CLng")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.trunc()))
}

pub(super) fn builtin_csng(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CSng")?;
    Ok(VBValue::Number(value_utils::to_arg_f64(&args[0])))
}

pub(super) fn builtin_ccur(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "CCur")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number((n * 10000.0).round() / 10000.0))
}

pub(super) fn builtin_hex(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Hex")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n == 0.0 {
        return Ok(VBValue::String("0".into()));
    }
    let n = n as i64;
    Ok(VBValue::String(format!("{:X}", n).into()))
}

pub(super) fn builtin_oct(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Oct")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n == 0.0 {
        return Ok(VBValue::String("0".into()));
    }
    let n = n as i64;
    Ok(VBValue::String(format!("{:o}", n).into()))
}

pub(super) fn builtin_isdate(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsDate")?;
    match &args[0] {
        VBValue::Number(_) => Ok(VBValue::Boolean(true)),
        VBValue::String(s) => Ok(VBValue::Boolean(super::datetime::try_parse_date(s).is_some())),
        _ => Ok(VBValue::Boolean(false)),
    }
}

pub(super) fn builtin_isobject(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "IsObject")?;
    Ok(VBValue::Boolean(matches!(args[0], VBValue::Object(_))))
}

pub(super) fn builtin_typename(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
    Ok(VBValue::String(name.to_string().into()))
}

pub(super) fn builtin_vartype(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
