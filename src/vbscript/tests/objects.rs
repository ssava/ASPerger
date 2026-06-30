use super::*;
    // ===== OBJECT / DICTIONARY / METHOD CALLS =====

    #[test]
    fn test_createobject_dictionary() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        let val = context.get_variable("dict");
        assert!(val.is_some());
        assert!(matches!(val.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_dictionary_method_call() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        interpreter
            .execute("dict.Add \"b\", \"Beta\"", &mut context)
            .unwrap();
    }

    #[test]
    fn test_dictionary_property_keys() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        interpreter
            .execute("dict.Add \"b\", \"Beta\"", &mut context)
            .unwrap();
        context.set_variable("keys", VBValue::Empty);
        interpreter
            .execute("keys = dict.Keys", &mut context)
            .unwrap();
        let keys = context.get_variable("keys");
        assert!(matches!(keys, Some(VBValue::Array(_, _))));
        if let Some(VBValue::Array(items, _)) = keys {
            assert_eq!(items.len(), 2);
        }
    }

    #[test]
    fn test_dictionary_indexed_access() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        context.set_variable("val", VBValue::Empty);
        interpreter
            .execute("val = dict(\"a\")", &mut context)
            .unwrap();
        assert_eq!(
            context.get_variable("val"),
            Some(&VBValue::String("Alpha".into()))
        );
    }

    #[test]
    fn test_for_each_with_dictionary_keys() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        interpreter
            .execute("dict.Add \"b\", \"Beta\"", &mut context)
            .unwrap();
        interpreter
            .execute("dict.Add \"g\", \"Gamma\"", &mut context)
            .unwrap();
        context.set_variable("result", VBValue::String("".into()));
        interpreter
            .execute(
                "For Each key In dict.Keys\n    result = result & key\nNext",
                &mut context,
            )
            .unwrap();
        let result = context.get_variable("result");
        assert!(result.is_some());
        // Keys may be in any order
        let s = match result.unwrap() {
            VBValue::String(s) => s.to_string(),
            _ => String::new(),
        };
        assert_eq!(s.len(), 3);
        assert!(s.contains('a'));
        assert!(s.contains('b'));
        assert!(s.contains('g'));
    }

    #[test]
    fn test_dictionary_count() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        interpreter
            .execute("dict.Add \"b\", \"Beta\"", &mut context)
            .unwrap();
        context.set_variable("cnt", VBValue::Empty);
        interpreter
            .execute("cnt = dict.Count", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("cnt"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_dictionary_exists() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        context.set_variable("found", VBValue::Empty);
        interpreter
            .execute("found = dict.Exists(\"a\")", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("found"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_method_call_no_args() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut context,
            )
            .unwrap();
        interpreter
            .execute("dict.Add \"a\", \"Alpha\"", &mut context)
            .unwrap();
        interpreter.execute("dict.RemoveAll", &mut context).unwrap();
        context.set_variable("cnt", VBValue::Empty);
        interpreter
            .execute("cnt = dict.Count", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("cnt"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_asp_index_page() {
        let content =
            fs::read_to_string("asp_files/index.asp").expect("Failed to read asp_files/index.asp");
        let parser = AspParser::new(content);
        let blocks = parser.parse();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let store = crate::vbscript::store::Store::new();
        let mut context = crate::vbscript::ExecutionContext::new();
        context.store = Some(Arc::clone(&store));
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let mut output = String::new();

        for block in blocks.iter() {
            match block {
                crate::asp::parser::AspBlock::Html(html) => {
                    output.push_str(html);
                }
                crate::asp::parser::AspBlock::Code(code, _code_line) => {
                    match interpreter.execute(code, &mut context) {
                        Ok(()) => {
                            output.push_str(&context.response.buffer);
                        }
                        Err(e) => {
                            output.push_str(&format!("<!-- Error: {} -->", e));
                        }
                    }
                    context.flush_response_buffer();
                }
                crate::asp::parser::AspBlock::Directive(_, _) => {
                    // Directives are informational; no runtime impact
                }
            }
        }

        // Find the summary header and extract counts
        let summary_prefix = "Summary: ";
        let summary_suffix = " passed";
        if let Some(start) = output.find(summary_prefix) {
            let after_prefix = &output[start + summary_prefix.len()..];
            if let Some(end) = after_prefix.find(summary_suffix) {
                let counts = &after_prefix[..end];
                if let Some(slash) = counts.find('/') {
                    let passed: i32 = counts[..slash].trim().parse().unwrap_or(-1);
                    let total: i32 = counts[slash + 1..].trim().parse().unwrap_or(-1);
                    assert_eq!(total, 29, "Expected 29 total tests, got {}", total);
                    assert_eq!(passed, 29, "Expected 29 passing tests, got {}. Check if unimplemented features changed", passed);
                    return;
                }
            }
        }
        panic!(
            "Summary not found in output. Output snippet (last 500 chars):\n{}\n---",
            &output[output.len().saturating_sub(500)..]
        );
    }

    // ===== ERR.RAISE =====

    #[test]
    fn test_err_raise_sets_err_state() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        // Err.Raise should produce an error that is caught in ResumeNext mode
        ctx.set_error_mode(crate::vbscript::execution_context::ErrorMode::ResumeNext);
        interp
            .execute(
                "On Error Resume Next\nErr.Raise 42, \"custom error\"",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.err_number, 42.0);
        assert_eq!(ctx.err_description, "custom error");
    }

    #[test]
    fn test_err_raise_without_args_sets_err() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_error_mode(crate::vbscript::execution_context::ErrorMode::ResumeNext);
        interp
            .execute("On Error Resume Next\nErr.Raise", &mut ctx)
            .unwrap();
        assert!(ctx.err_number != 0.0);
    }

    #[test]
    fn test_err_raise_min_number_only() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_error_mode(crate::vbscript::execution_context::ErrorMode::ResumeNext);
        interp
            .execute("On Error Resume Next\nErr.Raise 5", &mut ctx)
            .unwrap();
        assert_eq!(ctx.err_number, 5.0);
    }

    // ===== REGEXP OBJECT =====

    #[test]
    fn test_createobject_regexp() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set re = CreateObject(\"VBScript.RegExp\")", &mut ctx)
            .unwrap();
        let val = ctx.get_variable("re");
        assert!(val.is_some());
        assert!(matches!(val.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_regexp_test_true() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"hello\"\nresult = re.Test(\"hello world\")", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_regexp_test_false() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"hello\"\nresult = re.Test(\"goodbye world\")", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_regexp_replace() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"world\"\nre.Global = True\nresult = re.Replace(\"hello world world\", \"there\")", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hello there there".into()))
        );
    }

    #[test]
    fn test_regexp_replace_single() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"world\"\nresult = re.Replace(\"hello world world\", \"there\")", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hello there world".into()))
        );
    }

    #[test]
    fn test_regexp_execute() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"\\d+\"\nre.Global = True\nresult = re.Execute(\"abc 123 def 456\")", &mut ctx).unwrap();
        let result_val = ctx.get_variable("result").cloned();
        if let Some(VBValue::Array(arr, _)) = result_val {
            assert_eq!(arr.len(), 2);
            let match0_value = match &arr[0] {
                VBValue::Object(m) => m.get_property("Value", &mut ExecutionContext::new()).unwrap(),
                _ => panic!("Expected Match object at [0]"),
            };
            let match1_value = match &arr[1] {
                VBValue::Object(m) => m.get_property("Value", &mut ExecutionContext::new()).unwrap(),
                _ => panic!("Expected Match object at [1]"),
            };
            assert_eq!(match0_value, VBValue::String("123".into()));
            assert_eq!(match1_value, VBValue::String("456".into()));
        } else {
            panic!("Expected Array");
        }
    }

    #[test]
    fn test_regexp_ignorecase() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"hello\"\nre.IgnoreCase = True\nresult = re.Test(\"HELLO world\")", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_regexp_properties() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Set re = CreateObject(\"VBScript.RegExp\")\nre.Pattern = \"test\"\nre.IgnoreCase = True\nre.Global = True\np = re.Pattern\ni = re.IgnoreCase\ng = re.Global", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("p"),
            Some(&VBValue::String("test".into()))
        );
        assert_eq!(ctx.get_variable("i"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("g"), Some(&VBValue::Boolean(true)));
    }

    // ===== ADODB CONNECTION =====

    #[test]
    fn test_createobject_adodb_connection() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set conn = CreateObject(\"ADODB.Connection\")", &mut ctx)
            .unwrap();
        let val = ctx.get_variable("conn");
        assert!(val.is_some());
        assert!(matches!(val.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_adodb_connection_open_close() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set conn = CreateObject(\"ADODB.Connection\")", &mut ctx)
            .unwrap();
        interp.execute("conn.Open \"dsn=mydb\"", &mut ctx).unwrap();
        interp.execute("state = conn.State", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("state"), Some(&VBValue::Number(1.0)));
        interp.execute("conn.Close", &mut ctx).unwrap();
        interp.execute("state = conn.State", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("state"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_adodb_connection_execute_returns_recordset() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set conn = CreateObject(\"ADODB.Connection\")", &mut ctx)
            .unwrap();
        interp
            .execute("Set rs = conn.Execute(\"SELECT * FROM test\")", &mut ctx)
            .unwrap();
        let val = ctx.get_variable("rs");
        assert!(val.is_some());
        assert!(matches!(val.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_adodb_recordset_eof() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set conn = CreateObject(\"ADODB.Connection\")", &mut ctx)
            .unwrap();
        interp
            .execute("Set rs = conn.Execute(\"SELECT * FROM test\")", &mut ctx)
            .unwrap();
        interp.execute("eof = rs.EOF", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("eof"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_adodb_connection_connectionstring() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Set conn = CreateObject(\"ADODB.Connection\")", &mut ctx)
            .unwrap();
        interp
            .execute(
                "conn.ConnectionString = \"Provider=SQLOLEDB;Data Source=server\"",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("cs = conn.ConnectionString", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("cs"),
            Some(&VBValue::String(
                "Provider=SQLOLEDB;Data Source=server".to_string().into()
            ))
        );
    }

    // ===== LSET / RSET =====

    #[test]
    fn test_builtin_lset_pad() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = LSet(\"hi\", 5)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hi   ".into()))
        );
    }

    #[test]
    fn test_builtin_lset_truncate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = LSet(\"hello\", 3)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hel".into()))
        );
    }

    #[test]
    fn test_builtin_rset_pad() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = RSet(\"hi\", 5)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("   hi".into()))
        );
    }

    #[test]
    fn test_builtin_rset_truncate() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("result = RSet(\"hello\", 3)", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::String("hel".into()))
        );
    }
