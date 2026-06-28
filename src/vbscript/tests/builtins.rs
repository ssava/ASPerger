use super::*;
    // ===== BUILT-IN FUNCTIONS: BATCH 1 (SPLIT, JOIN, REPLACE, etc.) =====

    #[test]
    fn test_builtin_split_default_delimiter() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Split(\"a b c\")", &mut ctx)
            .unwrap();
        let arr = ctx.get_variable("result");
        assert!(matches!(arr, Some(VBValue::Array(..))));
        if let Some(VBValue::Array(a, _)) = arr {
            assert_eq!(a.len(), 3);
            assert_eq!(a[0], VBValue::String("a".to_string()));
            assert_eq!(a[1], VBValue::String("b".to_string()));
            assert_eq!(a[2], VBValue::String("c".to_string()));
        }
    }

    #[test]
    fn test_builtin_split_custom_delimiter() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Split(\"x,y,z\", \",\")", &mut ctx)
            .unwrap();
        if let Some(VBValue::Array(a, _)) = ctx.get_variable("result") {
            assert_eq!(a.len(), 3);
            assert_eq!(a[0], VBValue::String("x".to_string()));
            assert_eq!(a[1], VBValue::String("y".to_string()));
            assert_eq!(a[2], VBValue::String("z".to_string()));
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_builtin_split_with_count() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Split(\"a,b,c,d\", \",\", 2)", &mut ctx)
            .unwrap();
        if let Some(VBValue::Array(a, _)) = ctx.get_variable("result") {
            assert_eq!(a.len(), 2);
            assert_eq!(a[0], VBValue::String("a".to_string()));
            assert_eq!(a[1], VBValue::String("b,c,d".to_string()));
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_builtin_join() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Join(Array(\"a\", \"b\", \"c\"), \",\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("a,b,c".to_string()))
        );
    }

    #[test]
    fn test_builtin_join_default_delimiter() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Join(Array(\"x\", \"y\"))", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("x y".to_string()))
        );
    }

    #[test]
    fn test_builtin_replace() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "result = Replace(\"hello world world\", \"world\", \"there\")",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hello there there".to_string()))
        );
    }

    #[test]
    fn test_builtin_replace_with_count() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "result = Replace(\"a,b,c,d\", \",\", \"|\", 1, 2)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("a|b|c,d".to_string()))
        );
    }

    #[test]
    fn test_builtin_replace_with_start() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Replace(\"xxxyyyxxx\", \"x\", \"z\", 4)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("yyyzzz".to_string()))
        );
    }

    #[test]
    fn test_builtin_asc() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Asc(\"A\")", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(65.0)));
    }

    #[test]
    fn test_builtin_chr() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Chr(65)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("A".to_string()))
        );
    }

    #[test]
    fn test_builtin_ltrim() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = LTrim(\"  hello  \")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hello  ".to_string()))
        );
    }

    #[test]
    fn test_builtin_rtrim() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = RTrim(\"  hello  \")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("  hello".to_string()))
        );
    }

    #[test]
    fn test_builtin_space() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Space(5)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("     ".to_string()))
        );
    }

    #[test]
    fn test_builtin_string_number() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = String(3, 65)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("AAA".to_string()))
        );
    }

    #[test]
    fn test_builtin_string_char() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = String(5, \"*\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("*****".to_string()))
        );
    }

    #[test]
    fn test_builtin_strreverse() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrReverse(\"hello\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("olleh".to_string()))
        );
    }

    #[test]
    fn test_builtin_instrrev() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = InStrRev(\"abcabc\", \"ab\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_builtin_isnumeric_string() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = IsNumeric(\"123\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_builtin_isnumeric_non_numeric() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = IsNumeric(\"abc\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_builtin_isnumeric_number() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = IsNumeric(42)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_builtin_isarray_true() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = IsArray(Array(1, 2, 3))", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_builtin_isarray_false() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = IsArray(\"not an array\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(false)));
    }


    // ===== BUILT-IN FUNCTIONS: BATCH 2 (Date/Time, Math, Array, Type Conversion) =====

    #[test]
    fn test_builtin_now() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Now()", &mut ctx).unwrap();
        let val = ctx.get_variable("result");
        assert!(matches!(val, Some(VBValue::Number(_))));
        if let Some(VBValue::Number(n)) = val {
            assert!(*n > 0.0);
        }
    }

    #[test]
    fn test_builtin_date() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Date()", &mut ctx).unwrap();
        assert!(matches!(
            ctx.get_variable("result"),
            Some(VBValue::String(_))
        ));
        if let Some(VBValue::String(s)) = ctx.get_variable("result") {
            assert!(!s.is_empty());
            assert!(s.contains('/'));
        }
    }

    #[test]
    fn test_builtin_time() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Time()", &mut ctx).unwrap();
        assert!(matches!(
            ctx.get_variable("result"),
            Some(VBValue::String(_))
        ));
        if let Some(VBValue::String(s)) = ctx.get_variable("result") {
            assert!(!s.is_empty());
            assert!(s.contains(':'));
        }
    }

    #[test]
    fn test_builtin_year_month_day() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "d = DateSerial(2024, 6, 15)\ny = Year(d)\nm = Month(d)\ndy = Day(d)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("y"), Some(&VBValue::Number(2024.0)));
        assert_eq!(ctx.get_variable("m"), Some(&VBValue::Number(6.0)));
        assert_eq!(ctx.get_variable("dy"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_builtin_year_month_day_with_date_string() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "d = CDate(\"06/15/2024\")\ny = Year(d)\nm = Month(d)\ndy = Day(d)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("y"), Some(&VBValue::Number(2024.0)));
        assert_eq!(ctx.get_variable("m"), Some(&VBValue::Number(6.0)));
        assert_eq!(ctx.get_variable("dy"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_builtin_year_with_date_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("y = Year(\"2024-06-15\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("y"), Some(&VBValue::Number(2024.0)));
    }

    #[test]
    fn test_builtin_month_with_date_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("m = Month(\"2024-06-15\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("m"), Some(&VBValue::Number(6.0)));
    }

    #[test]
    fn test_builtin_day_with_date_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("dy = Day(\"2024-06-15\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("dy"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_builtin_hour_minute_second() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "t = TimeSerial(14, 30, 45)\nh = Hour(t)\nmi = Minute(t)\ns = Second(t)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("h"), Some(&VBValue::Number(14.0)));
        assert_eq!(ctx.get_variable("mi"), Some(&VBValue::Number(30.0)));
        assert_eq!(ctx.get_variable("s"), Some(&VBValue::Number(45.0)));
    }

    #[test]
    fn test_builtin_weekday() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        // 2024-01-07 is a Sunday
        interp
            .execute("d = DateSerial(2024, 1, 7)\nw = Weekday(d)", &mut ctx)
            .unwrap();
        // Sunday = 1
        assert_eq!(ctx.get_variable("w"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_builtin_weekday_with_firstday() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        // 2024-01-08 is a Monday
        interp
            .execute("d = DateSerial(2024, 1, 8)\nw = Weekday(d, 2)", &mut ctx)
            .unwrap();
        // With firstday=2 (Monday), Monday = 1
        assert_eq!(ctx.get_variable("w"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_builtin_weekdayname() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = WeekdayName(1)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("Sunday".to_string()))
        );
    }

    #[test]
    fn test_builtin_weekdayname_abbreviate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = WeekdayName(2, True)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("Mon".to_string()))
        );
    }

    #[test]
    fn test_builtin_monthname() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = MonthName(1)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("January".to_string()))
        );
    }

    #[test]
    fn test_builtin_monthname_abbreviate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = MonthName(2, True)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("Feb".to_string()))
        );
    }

    #[test]
    fn test_builtin_dateadd_days() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "d = DateSerial(2024, 1, 1)\nresult = DateAdd(\"d\", 10, d)",
                &mut ctx,
            )
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            assert_eq!(dt.day(), 11);
            assert_eq!(dt.month(), 1);
            assert_eq!(dt.year(), 2024);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_dateadd_months() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "d = DateSerial(2024, 1, 31)\nresult = DateAdd(\"m\", 1, d)",
                &mut ctx,
            )
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            // Jan 31 + 1 month = Feb 28 (or 29 in leap year; 2024 is leap)
            assert_eq!(dt.month(), 2);
            assert_eq!(dt.day(), 29);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_datediff_days() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("d1 = DateSerial(2024, 1, 1)\nd2 = DateSerial(2024, 1, 11)\nresult = DateDiff(\"d\", d1, d2)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(10.0)));
    }

    #[test]
    fn test_builtin_datediff_years() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("d1 = DateSerial(2020, 1, 1)\nd2 = DateSerial(2024, 1, 1)\nresult = DateDiff(\"yyyy\", d1, d2)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_builtin_dateserial() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = DateSerial(2024, 7, 4)", &mut ctx)
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            assert_eq!(dt.year(), 2024);
            assert_eq!(dt.month(), 7);
            assert_eq!(dt.day(), 4);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_datevalue() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = DateValue(\"2024-07-04\")", &mut ctx)
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            assert_eq!(dt.year(), 2024);
            assert_eq!(dt.month(), 7);
            assert_eq!(dt.day(), 4);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_timeserial() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = TimeSerial(10, 30, 0)", &mut ctx)
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            assert_eq!(dt.hour(), 10);
            assert_eq!(dt.minute(), 30);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_timevalue() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = TimeValue(\"14:30:00\")", &mut ctx)
            .unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            let dt = crate::vbscript::builtins::ole_auto_to_datetime(*n).unwrap();
            assert_eq!(dt.hour(), 14);
            assert_eq!(dt.minute(), 30);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_timer() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Timer()", &mut ctx).unwrap();
        let val = ctx.get_variable("result");
        assert!(matches!(val, Some(VBValue::Number(_))));
        if let Some(VBValue::Number(n)) = val {
            assert!(*n >= 0.0 && *n < 86400.0);
        }
    }

    #[test]
    fn test_builtin_formatdatetime() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "d = DateSerial(2024, 7, 4)\nresult = FormatDateTime(d, 2)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("07/04/2024".to_string()))
        );
    }

    #[test]
    fn test_builtin_int_vs_fix_negative() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("i = Int(-3.1)\nf = Fix(-3.1)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("i"), Some(&VBValue::Number(-4.0)));
        assert_eq!(ctx.get_variable("f"), Some(&VBValue::Number(-3.0)));
    }

    #[test]
    #[allow(clippy::approx_constant)]
    fn test_builtin_round() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = Round(3.14159, 2)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(3.14)));
    }

    #[test]
    fn test_builtin_sgn() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("r1 = Sgn(5)\nr2 = Sgn(0)\nr3 = Sgn(-3)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Number(1.0)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Number(0.0)));
        assert_eq!(ctx.get_variable("r3"), Some(&VBValue::Number(-1.0)));
    }

    #[test]
    fn test_builtin_sqr() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Sqr(9)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(3.0)));
    }

    #[test]
    fn test_builtin_sqr_negative_error() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        let result = interp.execute("result = Sqr(-1)", &mut ctx);
        assert!(result.is_err());
    }

    #[test]
    fn test_builtin_ubound_array() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("a = Array(10, 20, 30)\nresult = UBound(a)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_builtin_lbound_array() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("a = Array(1, 2, 3)\nresult = LBound(a)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_builtin_cbool() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "r1 = CBool(1)\nr2 = CBool(0)\nr3 = CBool(\"True\")\nr4 = CBool(\"False\")",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Boolean(false)));
        assert_eq!(ctx.get_variable("r3"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("r4"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_builtin_cbyte() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = CByte(42)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(42.0)));
    }

    #[test]
    fn test_builtin_cbyte_overflow() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        let result = interp.execute("result = CByte(300)", &mut ctx);
        assert!(result.is_err());
    }

    #[test]
    fn test_builtin_cdate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = CDate(\"2024-07-04\")", &mut ctx)
            .unwrap();
        assert!(matches!(
            ctx.get_variable("result"),
            Some(VBValue::Number(_))
        ));
    }

    #[test]
    fn test_builtin_cdbl() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = CDbl(42)", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(42.0)));
    }

    #[test]
    fn test_builtin_clng() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("r1 = CLng(3.14)\nr2 = CLng(-3.9)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Number(3.0)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Number(-3.0)));
    }

    #[test]
    fn test_builtin_hex() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("r1 = Hex(255)\nr2 = Hex(0)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("r1"),
            Some(&VBValue::String("FF".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r2"),
            Some(&VBValue::String("0".to_string()))
        );
    }

    #[test]
    fn test_builtin_oct() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Oct(8)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("10".to_string()))
        );
    }

    #[test]
    fn test_builtin_isdate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "r1 = IsDate(\"2024-01-15\")\nr2 = IsDate(\"not a date\")",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_builtin_isobject() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set dict = CreateObject(\"Scripting.Dictionary\")\nr1 = IsObject(dict)\nr2 = IsObject(\"string\")", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_builtin_typename() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute(
            "r1 = TypeName(\"hello\")\nr2 = TypeName(42)\nr3 = TypeName(123456)\nr4 = TypeName(3.14)\nr5 = TypeName(True)\nr6 = TypeName(Null)\nr7 = TypeName(Empty)\nr8 = TypeName(Array(1,2))",
            &mut ctx,
        ).unwrap();
        assert_eq!(
            ctx.get_variable("r1"),
            Some(&VBValue::String("String".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r2"),
            Some(&VBValue::String("Integer".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r3"),
            Some(&VBValue::String("Long".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r4"),
            Some(&VBValue::String("Double".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r5"),
            Some(&VBValue::String("Boolean".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r6"),
            Some(&VBValue::String("Null".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r7"),
            Some(&VBValue::String("Empty".to_string()))
        );
        assert_eq!(
            ctx.get_variable("r8"),
            Some(&VBValue::String("Array".to_string()))
        );
    }

    #[test]
    fn test_builtin_vartype() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("r1 = VarType(\"hello\")\nr2 = VarType(42)\nr3 = VarType(True)\nr4 = VarType(Null)\nr5 = VarType(Empty)\nr6 = VarType(Array(1,2))", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Number(8.0))); // vbString
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Number(2.0))); // vbInteger
        assert_eq!(ctx.get_variable("r3"), Some(&VBValue::Number(11.0))); // vbBoolean
        assert_eq!(ctx.get_variable("r4"), Some(&VBValue::Number(1.0))); // vbNull
        assert_eq!(ctx.get_variable("r5"), Some(&VBValue::Number(0.0))); // vbEmpty
        assert_eq!(ctx.get_variable("r6"), Some(&VBValue::Number(8204.0))); // vbArray + vbVariant
    }

    #[test]
    fn test_builtin_rnd_range() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("result = Rnd()", &mut ctx).unwrap();
        if let Some(VBValue::Number(n)) = ctx.get_variable("result") {
            assert!(*n >= 0.0 && *n < 1.0);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_builtin_filter_include() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("a = Array(\"apple\", \"banana\", \"apricot\", \"cherry\")\nresult = Filter(a, \"ap\")", &mut ctx).unwrap();
        if let Some(VBValue::Array(arr, _)) = ctx.get_variable("result") {
            assert_eq!(arr.len(), 2);
            assert_eq!(arr[0], VBValue::String("apple".to_string()));
            assert_eq!(arr[1], VBValue::String("apricot".to_string()));
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_builtin_filter_exclude() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "a = Array(\"apple\", \"banana\", \"apricot\")\nresult = Filter(a, \"ap\", False)",
                &mut ctx,
            )
            .unwrap();
        if let Some(VBValue::Array(arr, _)) = ctx.get_variable("result") {
            assert_eq!(arr.len(), 1);
            assert_eq!(arr[0], VBValue::String("banana".to_string()));
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_builtin_isarray_new() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("r1 = IsArray(Array(1,2))\nr2 = IsArray(\"not\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("r1"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("r2"), Some(&VBValue::Boolean(false)));
    }


    // ===== STRCOMP =====

    #[test]
    fn test_builtin_strcomp_equal() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrComp(\"hello\", \"hello\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_builtin_strcomp_less() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrComp(\"abc\", \"xyz\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(-1.0)));
    }

    #[test]
    fn test_builtin_strcomp_greater() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrComp(\"xyz\", \"abc\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_builtin_strcomp_textmode() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrComp(\"HELLO\", \"hello\", 1)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_builtin_strcomp_binarymode() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = StrComp(\"HELLO\", \"hello\", 0)", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(-1.0)));
    }

    // ===== FORMATNUMBER =====

    #[test]
    fn test_builtin_formatnumber_basic() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = FormatNumber(1234.567, 2)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("1,234.57".to_string()))
        );
    }

    #[test]
    fn test_builtin_formatnumber_no_decimal() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = FormatNumber(1234, 0)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("1,234".to_string()))
        );
    }

    // ===== FORMATCURRENCY =====

    #[test]
    fn test_builtin_formatcurrency_basic() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = FormatCurrency(1234.5)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("$1,234.50".to_string()))
        );
    }

    // ===== FORMATPERCENT =====

    #[test]
    fn test_builtin_formatpercent_basic() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = FormatPercent(0.1234, 1)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("12.3%".to_string()))
        );
    }
