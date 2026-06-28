//! Built-in VBScript functions: string, math, date/time, type conversion,
//! array, and miscellaneous operations dispatched by name.

mod string;
mod datetime;
mod math;
mod conv_misc;

#[cfg_attr(not(test), allow(unused_imports))]
pub(crate) use datetime::{datetime_to_ole_auto, ole_auto_to_datetime, try_parse_date};

macro_rules! builtins {
    ($name:ident, $args:ident, $($entry:literal => $func:ident),* $(,)?) => {
        match $name.to_uppercase().as_str() {
            $($entry => $func(&$args),)*
            _ => Err(crate::vbscript::vbs_error::VBSErrorType::NotImplementedError
                .into_error(format!("Function '{}' is not implemented", $name))),
        }
    };
}

/// Dispatch a built-in VBScript function call by name.
pub fn call_builtin(name: &str, args: Vec<VBValue>) -> Result<VBValue, VBSError> {
    use self::conv_misc::*;
    use self::datetime::*;
    use self::math::*;
    use self::string::*;

    builtins!(name, args,
        "ARRAY" => builtin_array,
        "FETCHURL" => builtin_fetchurl,
        "CREATEOBJECT" => builtin_createobject,
        "LEN" => builtin_len,
        "UCASE" => builtin_ucase,
        "LCASE" => builtin_lcase,
        "MID" => builtin_mid,
        "LEFT" => builtin_left,
        "RIGHT" => builtin_right,
        "TRIM" => builtin_trim,
        "CINT" => builtin_cint,
        "CSTR" => builtin_cstr,
        "ABS" => builtin_abs,
        "ISNULL" => builtin_isnull,
        "ISEMPTY" => builtin_isempty,
        "INSTR" => builtin_instr,
        "SPLIT" => builtin_split,
        "JOIN" => builtin_join,
        "REPLACE" => builtin_replace,
        "ASC" => builtin_asc,
        "CHR" => builtin_chr,
        "LTRIM" => builtin_ltrim,
        "RTRIM" => builtin_rtrim,
        "SPACE" => builtin_space,
        "STRING" => builtin_string,
        "STRREVERSE" => builtin_strreverse,
        "INSTRREV" => builtin_instrrev,
        "ISNUMERIC" => builtin_isnumeric,
        "ISARRAY" => builtin_isarray,
        "NOW" => builtin_now,
        "DATE" => builtin_date,
        "TIME" => builtin_time,
        "YEAR" => builtin_year,
        "MONTH" => builtin_month,
        "DAY" => builtin_day,
        "HOUR" => builtin_hour,
        "MINUTE" => builtin_minute,
        "SECOND" => builtin_second,
        "WEEKDAY" => builtin_weekday,
        "WEEKDAYNAME" => builtin_weekdayname,
        "MONTHNAME" => builtin_monthname,
        "DATEADD" => builtin_dateadd,
        "DATEDIFF" => builtin_datediff,
        "DATEPART" => builtin_datepart,
        "DATESERIAL" => builtin_dateserial,
        "DATEVALUE" => builtin_datevalue,
        "TIMESERIAL" => builtin_timeserial,
        "TIMEVALUE" => builtin_timevalue,
        "TIMER" => builtin_timer,
        "FORMATDATETIME" => builtin_formatdatetime,
        "RND" => builtin_rnd,
        "RANDOMIZE" => builtin_randomize,
        "INT" => builtin_int,
        "FIX" => builtin_fix,
        "ROUND" => builtin_round,
        "SGN" => builtin_sgn,
        "SQR" => builtin_sqr,
        "UBOUND" => builtin_ubound,
        "LBOUND" => builtin_lbound,
        "FILTER" => builtin_filter,
        "CBOOL" => builtin_cbool,
        "CBYTE" => builtin_cbyte,
        "CDATE" => builtin_cdate,
        "CDBL" => builtin_cdbl,
        "CLNG" => builtin_clng,
        "CSNG" => builtin_csng,
        "CCUR" => builtin_ccur,
        "HEX" => builtin_hex,
        "OCT" => builtin_oct,
        "ISDATE" => builtin_isdate,
        "ISOBJECT" => builtin_isobject,
        "TYPENAME" => builtin_typename,
        "VARTYPE" => builtin_vartype,
        "STRCOMP" => builtin_strcomp,
        "FORMATCURRENCY" => builtin_formatcurrency,
        "FORMATNUMBER" => builtin_formatnumber,
        "FORMATPERCENT" => builtin_formatpercent,
        "LSET" => builtin_lset,
        "RSET" => builtin_rset,
    )
}

use crate::vbscript::value::VBValue;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};

pub(crate) fn expect_arg_count(args: &[VBValue], expected: usize, name: &str) -> Result<(), VBSError> {
    if args.len() != expected {
        return Err(VBSErrorType::ValueError.into_error(format!(
            "{} requires {} argument(s), got {}",
            name, expected, args.len()
        )));
    }
    Ok(())
}

pub(crate) fn expect_min_args(args: &[VBValue], min: usize, name: &str) -> Result<(), VBSError> {
    if args.len() < min {
        return Err(VBSErrorType::ValueError.into_error(format!(
            "{} requires at least {} argument(s), got {}",
            name, min, args.len()
        )));
    }
    Ok(())
}

fn builtin_array(args: &[VBValue]) -> Result<VBValue, VBSError> {
    Ok(VBValue::Array(std::sync::Arc::new(args.to_vec()), vec![]))
}
