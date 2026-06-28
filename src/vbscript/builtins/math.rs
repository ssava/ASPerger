use super::expect_arg_count;
use super::expect_min_args;
use crate::vbscript::value::VBValue;
use crate::vbscript::value_utils;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use std::sync::Mutex;

static RNG_STATE: Mutex<u32> = Mutex::new(0u32);
static RNG_LAST: Mutex<f64> = Mutex::new(0.0f64);

fn rnd_next(state: &mut u32) -> f64 {
    *state = state.wrapping_mul(1103515245).wrapping_add(12345) & 0x7fffffff;
    *state as f64 / 2147483648.0
}

pub(super) fn builtin_abs(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Abs")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.abs()))
}

pub(super) fn builtin_rnd(args: &[VBValue]) -> Result<VBValue, VBSError> {
    let seed = if args.is_empty() {
        None
    } else {
        Some(value_utils::to_arg_f64(&args[0]))
    };
    let mut state = RNG_STATE.lock().unwrap_or_else(|e| e.into_inner());
    let mut last = RNG_LAST.lock().unwrap_or_else(|e| e.into_inner());
    match seed {
        Some(n) if n < 0.0 => {
            *state = ((-n) as u32) & 0x7fffffff;
            let val = rnd_next(&mut state);
            *last = val;
            Ok(VBValue::Number(val))
        }
        Some(0.0) => Ok(VBValue::Number(*last)),
        _ => {
            let val = rnd_next(&mut state);
            *last = val;
            Ok(VBValue::Number(val))
        }
    }
}

pub(super) fn builtin_randomize(args: &[VBValue]) -> Result<VBValue, VBSError> {
    let seed = if args.is_empty() {
        let now = chrono::Local::now().naive_local();
        now.and_utc().timestamp_subsec_nanos()
    } else {
        value_utils::to_arg_f64(&args[0]) as u32
    };
    *RNG_STATE.lock().unwrap_or_else(|e| e.into_inner()) = seed & 0x7fffffff;
    Ok(VBValue::Null)
}

pub(super) fn builtin_int(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Int")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.floor()))
}

pub(super) fn builtin_fix(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Fix")?;
    let n = value_utils::to_arg_f64(&args[0]);
    Ok(VBValue::Number(n.trunc()))
}

pub(super) fn builtin_round(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_sgn(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_sqr(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Sqr")?;
    let n = value_utils::to_arg_f64(&args[0]);
    if n < 0.0 {
        return Err(VBSErrorType::RuntimeError
            .into_error("Cannot calculate square root of a negative number".to_string()));
    }
    Ok(VBValue::Number(n.sqrt()))
}
