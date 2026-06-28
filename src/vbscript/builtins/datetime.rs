use super::expect_arg_count;
use super::expect_min_args;
use crate::vbscript::value::VBValue;
use crate::vbscript::value_utils;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use chrono::{Datelike, Timelike};

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

pub(super) fn builtin_now(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::Number(datetime_to_ole_auto(now)))
}

pub(super) fn builtin_date(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::String(now.format("%m/%d/%Y").to_string()))
}

pub(super) fn builtin_time(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    Ok(VBValue::String(now.format("%H:%M:%S").to_string()))
}

pub(super) fn builtin_year(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Year")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.year() as f64))
}

pub(super) fn builtin_month(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Month")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.month() as f64))
}

pub(super) fn builtin_day(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Day")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid date".to_string()))?;
    Ok(VBValue::Number(dt.day() as f64))
}

pub(super) fn builtin_hour(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Hour")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.hour() as f64))
}

pub(super) fn builtin_minute(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Minute")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.minute() as f64))
}

pub(super) fn builtin_second(args: &[VBValue]) -> Result<VBValue, VBSError> {
    expect_arg_count(args, 1, "Second")?;
    let dt = value_to_datetime(&args[0])
        .ok_or_else(|| VBSErrorType::RuntimeError.into_error("Invalid time".to_string()))?;
    Ok(VBValue::Number(dt.second() as f64))
}

pub(super) fn builtin_weekday(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_weekdayname(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_monthname(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_dateadd(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_datediff(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
        "y" | "d" => dt2.signed_duration_since(dt1).num_days() as f64,
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
        "ww" => dt2.signed_duration_since(dt1).num_days() as f64 / 7.0,
        "h" => dt2.signed_duration_since(dt1).num_hours() as f64,
        "n" => dt2.signed_duration_since(dt1).num_minutes() as f64,
        "s" => dt2.signed_duration_since(dt1).num_seconds() as f64,
        _ => {
            return Err(
                VBSErrorType::RuntimeError.into_error(format!("Invalid interval '{}'", interval))
            )
        }
    };
    Ok(VBValue::Number(result))
}

pub(super) fn builtin_datepart(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_dateserial(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_datevalue(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_timeserial(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_timevalue(args: &[VBValue]) -> Result<VBValue, VBSError> {
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

pub(super) fn builtin_timer(_args: &[VBValue]) -> Result<VBValue, VBSError> {
    let now = chrono::Local::now().naive_local();
    let seconds = now.num_seconds_from_midnight() as f64;
    let nanos = now.and_utc().timestamp_subsec_nanos() as f64;
    Ok(VBValue::Number(seconds + nanos / 1_000_000_000.0))
}

pub(super) fn builtin_formatdatetime(args: &[VBValue]) -> Result<VBValue, VBSError> {
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
