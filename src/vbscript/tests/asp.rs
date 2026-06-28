use super::*;
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};

    // ===== ASP PARSER =====

    #[test]
    fn test_asp_parser_splits_html_and_code() {
        let parser = AspParser::new("<html><%Dim x%></html>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 3);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Html(h) => assert_eq!(h, "<html>"),
            _ => panic!("Expected Html block"),
        }
        match &blocks[1] {
            crate::asp::parser::AspBlock::Code(c, _) => assert_eq!(c, "Dim x"),
            _ => panic!("Expected Code block"),
        }
    }

    #[test]
    fn test_asp_parser_only_html() {
        let parser = AspParser::new("<html><body>Hello</body></html>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Html(h) => {
                assert_eq!(h, "<html><body>Hello</body></html>")
            }
            _ => panic!("Expected Html block"),
        }
    }

    #[test]
    fn test_asp_parser_only_code() {
        let parser = AspParser::new("<%x = 1%>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Code(c, _) => assert_eq!(c, "x = 1"),
            _ => panic!("Expected Code block"),
        }
    }

    #[test]
    fn test_asp_parser_multiple_code_blocks() {
        let parser = AspParser::new("<%a = 1%><%b = 2%>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 2);
    }

    #[test]
    fn test_asp_parser_leading_trailing_html() {
        let parser = AspParser::new("before<%code%>after".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 3);
    }

    // ===== FOR EACH =====

    #[test]
    fn test_for_each_basic() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable(
            "items",
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(10.0),
                VBValue::Number(20.0),
                VBValue::Number(30.0),
            ]), vec![]),
        );
        context.set_variable("sum", VBValue::Number(0.0));
        interpreter
            .execute("For Each x In items\n    sum = sum + x\nNext", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("sum"), Some(&VBValue::Number(60.0)));
    }

    #[test]
    fn test_for_each_empty_array() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("items", VBValue::Array(std::sync::Arc::new(vec![]), vec![]));
        context.set_variable("flag", VBValue::Boolean(false));
        interpreter
            .execute("For Each x In items\n    flag = True\nNext", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("flag"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_for_each_string_array() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable(
            "items",
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::String("a".to_string()),
                VBValue::String("b".to_string()),
                VBValue::String("c".to_string()),
            ]), vec![]),
        );
        context.set_variable("result", VBValue::String("".to_string()));
        interpreter
            .execute(
                "For Each x In items\n    result = result & x\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(
            context.get_variable("result"),
            Some(&VBValue::String("abc".to_string()))
        );
    }

    #[test]
    fn test_for_each_non_array_error() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("For Each x In 42\nNext", &mut context);
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.message.contains("Object doesn't support"));
    }

    #[test]
    fn test_for_each_nested() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable(
            "outer",
            VBValue::Array(std::sync::Arc::new(vec![
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(1.0),
                VBValue::Number(2.0),
            ]), vec![]),
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(3.0),
                VBValue::Number(4.0),
            ]), vec![]),
            ]), vec![]),
        );
        context.set_variable("sum", VBValue::Number(0.0));
        interpreter.execute(
            "For Each row In outer\n    For Each col In row\n        sum = sum + col\n    Next\nNext",
            &mut context
        ).unwrap();
        assert_eq!(context.get_variable("sum"), Some(&VBValue::Number(10.0)));
    }

    #[test]
    fn test_for_each_modifies_element() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable(
            "items",
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(1.0),
                VBValue::Number(2.0),
                VBValue::Number(3.0),
            ]), vec![]),
        );
        context.set_variable("sum", VBValue::Number(0.0));
        interpreter
            .execute(
                "For Each x In items\n    sum = sum + x\n    x = 999\nNext",
                &mut context,
            )
            .unwrap();
        // x is overwritten each iteration, so sum is still 1+2+3=6
        assert_eq!(context.get_variable("sum"), Some(&VBValue::Number(6.0)));
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(999.0)));
    }

    #[test]
    fn test_for_each_with_for() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable(
            "items",
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(2.0),
                VBValue::Number(3.0),
                VBValue::Number(4.0),
            ]), vec![]),
        );
        context.set_variable("total", VBValue::Number(0.0));
        interpreter.execute(
            "For Each x In items\n    For i = 1 To x\n        total = total + i\n    Next\nNext",
            &mut context
        ).unwrap();
        // For x=2: 1+2=3. For x=3: 1+2+3=6. For x=4: 1+2+3+4=10. total=3+6+10=19
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(19.0)));
    }

    // ===== FUNCTION CALLS =====

    #[test]
    fn test_function_call_array() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("result", VBValue::String("".to_string()));
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = Array(10, 20, 30)", &mut context)
            .unwrap();
        assert_eq!(
            context.get_variable("result"),
            Some(&VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(10.0),
                VBValue::Number(20.0),
                VBValue::Number(30.0),
            ]), vec![]))
        );
    }

    #[test]
    fn test_function_call_array_in_for_each() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("sum", VBValue::Number(0.0));
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "For Each x In Array(1, 2, 3, 4, 5)\n    sum = sum + x\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("sum"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_function_call_len() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = Len(\"hello\")", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("result"), Some(&VBValue::Number(5.0)));
    }

    #[test]
    fn test_function_call_ucase() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = UCase(\"hello\")", &mut context)
            .unwrap();
        assert_eq!(
            context.get_variable("result"),
            Some(&VBValue::String("HELLO".to_string()))
        );
    }

    #[test]
    fn test_function_call_mid() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = Mid(\"hello\", 2, 3)", &mut context)
            .unwrap();
        assert_eq!(
            context.get_variable("result"),
            Some(&VBValue::String("ell".to_string()))
        );
    }

    #[test]
    fn test_function_call_unknown() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("result = UnknownFunc(42)", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_function_call_empty_args() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = Array()", &mut context)
            .unwrap();
        assert_eq!(
            context.get_variable("result"),
            Some(&VBValue::Array(std::sync::Arc::new(vec![]), vec![]))
        );
    }

    #[test]
    fn test_function_call_in_expression() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("x", VBValue::Number(3.0));
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("result = Len(\"abc\") + x", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("result"), Some(&VBValue::Number(6.0)));
    }


    // ===== ASP INTRINSIC OBJECTS =====

    #[test]
    fn test_asp_request_querystring() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request
            .params
            .insert("name".to_string(), "John".to_string());
        ctx.request
            .params
            .insert("age".to_string(), "30".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim n\nn = Request.QueryString(\"name\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("n"),
            Some(&VBValue::String("John".to_string()))
        );
    }

    #[test]
    fn test_asp_request_querystring_missing() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim n\nn = Request.QueryString(\"missing\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("n"),
            Some(&VBValue::String("".to_string()))
        );
    }

    #[test]
    fn test_asp_request_querystring_count() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request.params.insert("a".to_string(), "1".to_string());
        ctx.request.params.insert("b".to_string(), "2".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim c\nc = Request.QueryString.Count", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_asp_request_form() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request
            .form
            .insert("username".to_string(), "admin".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim u\nu = Request.Form(\"username\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("u"),
            Some(&VBValue::String("admin".to_string()))
        );
    }

    #[test]
    fn test_asp_request_form_count() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request.form.insert("x".to_string(), "1".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim c\nc = Request.Form.Count", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_asp_request_servervariables() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request
            .headers
            .insert("user-agent".to_string(), "ASPerger/1.0".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Dim u\nu = Request.ServerVariables(\"USER-AGENT\")",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("u"),
            Some(&VBValue::String("ASPerger/1.0".to_string()))
        );
    }

    #[test]
    fn test_asp_request_servervariables_count() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request
            .headers
            .insert("host".to_string(), "localhost".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim c\nc = Request.ServerVariables.Count", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_asp_request_cookies() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.request
            .cookies
            .insert("theme".to_string(), "dark".to_string());
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim t\nt = Request.Cookies(\"theme\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("t"),
            Some(&VBValue::String("dark".to_string()))
        );
    }

    #[test]
    fn test_asp_response_status_property() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("x = Response.Status", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("x"),
            Some(&VBValue::String("200 OK".to_string()))
        );
    }

    #[test]
    fn test_asp_response_contenttype_property() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("x = Response.ContentType", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("x"),
            Some(&VBValue::String("text/html".to_string()))
        );
    }

    #[test]
    fn test_asp_response_buffer_property() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("x = Response.Buffer", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_asp_response_write_syntax_shortcut() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Response.Write \"hello from ASP\"", &mut ctx)
            .unwrap();
        assert_eq!(ctx.response.buffer, "hello from ASP");
    }

    #[test]
    fn test_asp_response_write_with_variable() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        ctx.set_variable("name", VBValue::String("World".to_string()));
        let interp = VBScriptInterpreter;
        interp.execute("Response.Write name", &mut ctx).unwrap();
        assert_eq!(ctx.response.buffer, "World");
    }

    #[test]
    fn test_asp_response_binarywrite_with_array() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("data = Array(72, 101, 108, 108, 111): Response.BinaryWrite data", &mut ctx)
            .unwrap();
        assert_eq!(ctx.response.binary_buffer, vec![72, 101, 108, 108, 111]);
    }

    #[test]
    fn test_asp_response_binarywrite_with_string() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Response.BinaryWrite \"Hello\"", &mut ctx)
            .unwrap();
        assert_eq!(ctx.response.binary_buffer, b"Hello");
    }

    #[test]
    fn test_asp_response_ended() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        // Manually simulate Response.End - this would be called through the object
        ctx.response.ended = true;
        let interp = VBScriptInterpreter;
        interp.execute("x = 42", &mut ctx).unwrap();
        // x should not be set because response_ended prevents execution
        assert_eq!(ctx.get_variable("x"), None);
    }

    #[test]
    fn test_asp_session_set_and_get() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        ctx.session.id = "test-session-001".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        // Session("key") = value syntax goes through indexed_set
        // The test uses the property syntax
        interp
            .execute("Session(\"username\") = \"Alice\"", &mut ctx)
            .unwrap();
        // Check stored value via local store
        let session_val = {
            let data = store.lock_sessions();
            data.get("TEST-SESSION-001")
                .and_then(|d| d.get("USERNAME"))
                .cloned()
        };
        assert_eq!(session_val, Some(VBValue::String("Alice".to_string())));
    }

    #[test]
    fn test_asp_session_sessionid() {
        let mut ctx = ExecutionContext::new();
        ctx.session.id = "MY-SESSION".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim s\ns = Session.SessionID", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("s"),
            Some(&VBValue::String("MY-SESSION".to_string()))
        );
    }

    #[test]
    fn test_asp_session_timeout() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim t\nt = Session.Timeout", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("t"), Some(&VBValue::Number(20.0)));
    }

    #[test]
    fn test_asp_session_contents_count() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        ctx.session.id = "test-contents".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Session(\"k1\") = \"v1\"\nSession(\"k2\") = \"v2\"",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("Dim c\nc = Session.Contents.Count", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_asp_session_contents_key_item() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        ctx.session.id = "test-contents-key-item".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Session(\"k1\") = \"v1\"\nSession(\"k2\") = \"v2\"",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("Dim k1, k2, v1, v2\nk1 = Session.Contents.Key(1)\nk2 = Session.Contents.Key(2)\nv1 = Session.Contents.Item(1)\nv2 = Session.Contents.Item(2)", &mut ctx)
            .unwrap();
        let k1 = ctx.get_variable("k1").unwrap().clone();
        let k2 = ctx.get_variable("k2").unwrap().clone();
        let v1 = ctx.get_variable("v1").unwrap().clone();
        let v2 = ctx.get_variable("v2").unwrap().clone();
        assert_ne!(k1, k2, "keys should be distinct");
        assert_ne!(v1, v2, "values should be distinct");
        match (&k1, &k2) {
            (VBValue::String(a), VBValue::String(b)) => {
                let expected = ["K1", "K2"];
                assert!(expected.contains(&a.as_str()), "k1={:?} not in [K1,K2]", a);
                assert!(expected.contains(&b.as_str()), "k2={:?} not in [K1,K2]", b);
            }
            _ => panic!("Expected String keys, got {:?} {:?}", k1, k2),
        }
        assert!(
            v1 == VBValue::String("v1".to_string()) || v1 == VBValue::String("v2".to_string()),
            "v1={:?} not in [v1,v2]",
            v1
        );
        assert!(
            v2 == VBValue::String("v1".to_string()) || v2 == VBValue::String("v2".to_string()),
            "v2={:?} not in [v1,v2]",
            v2
        );
    }

    #[test]
    fn test_asp_session_contents_indexed_get() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        ctx.session.id = "test-contents-indexed".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Session(\"k1\") = \"v1\"\nSession(\"k2\") = \"v2\"",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("Dim v\nv = Session.Contents(\"k1\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("v"),
            Some(&VBValue::String("v1".to_string()))
        );
    }

    #[test]
    fn test_asp_session_abandon() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        ctx.session.id = "abandon-test".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Session(\"data\") = \"keep me\"", &mut ctx)
            .unwrap();
        interp.execute("Session.Abandon", &mut ctx).unwrap();
        let data = store.lock_sessions();
        assert!(!data.contains_key("ABANDON-TEST"));
    }

    #[test]
    fn test_asp_server_htmlencode() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Dim r\nr = Server.HTMLEncode(\"<b>bold</b> & 'quotes'\")",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("r"),
            Some(&VBValue::String(
                "&lt;b&gt;bold&lt;/b&gt; &amp; &#39;quotes&#39;".to_string()
            ))
        );
    }

    #[test]
    fn test_asp_server_urlencode() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim r\nr = Server.URLEncode(\"hello world\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("r"),
            Some(&VBValue::String("hello+world".to_string()))
        );
    }

    #[test]
    fn test_asp_server_urlencode_special_chars() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim r\nr = Server.URLEncode(\"a/b?c\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("r"),
            Some(&VBValue::String("a%2Fb%3Fc".to_string()))
        );
    }

    #[test]
    fn test_asp_server_urlpathencode() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim r\nr = Server.URLPathEncode(\"a/b c\")", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("r"),
            Some(&VBValue::String("a/b+c".to_string()))
        );
    }

    #[test]
    fn test_asp_server_scripttimeout() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim t\nt = Server.ScriptTimeout", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("t"), Some(&VBValue::Number(90.0)));
    }

    #[test]
    fn test_asp_server_scriptpath() {
        let mut ctx = ExecutionContext::new();
        ctx.script_path = "/home/site/index.asp".to_string();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim p\np = Server.ScriptPath", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("p"),
            Some(&VBValue::String("/home/site/index.asp".to_string()))
        );
    }

    #[test]
    fn test_asp_server_createobject_dictionary() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set dict = Server.CreateObject(\"Scripting.Dictionary\")",
                &mut ctx,
            )
            .unwrap();
        let val = ctx.get_variable("dict");
        assert!(val.is_some());
        assert!(matches!(val.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_asp_server_mappath() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim p\np = Server.MapPath(\"/test.txt\")", &mut ctx)
            .unwrap();
        let val = ctx.get_variable("p");
        assert!(val.is_some());
        if let Some(VBValue::String(s)) = val {
            assert!(s.contains("test.txt"));
        }
    }

    #[test]
    fn test_asp_application_set_and_get() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Application.Lock\nApplication(\"counter\") = 42\nApplication.Unlock",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("Dim c\nc = Application(\"counter\")", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(42.0)));
    }

    #[test]
    fn test_asp_application_contents_count() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Application(\"ac_cnt_k1\") = 1\nApplication(\"ac_cnt_k2\") = 2",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("Dim c\nc = Application.Contents.Count", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("c"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_asp_application_lock_unlock() {
        let store = crate::vbscript::store::Store::new();
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(Arc::clone(&store));
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Application.Lock\nApplication(\"key\") = \"val\"\nApplication.Unlock",
                &mut ctx,
            )
            .unwrap();
        let val = {
            let data = store.lock_apps();
            data.get("KEY").cloned()
        };
        assert_eq!(val, Some(VBValue::String("val".to_string())));
    }

    #[test]
    fn test_asp_request_totalbytes() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim t\nt = Request.TotalBytes", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("t"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_asp_response_expires() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Dim e\ne = Response.Expires", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("e"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_asp_response_cookies_set() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        // The cookies collection needs to be accessed via Response.Cookies
        let _ = interp.execute("Response.Cookies(\"test\") = \"value\"", &mut ctx);
        // Just ensure no crash for now (property set on cookies)
    }

    #[test]
    fn test_asp_response_cookies_set_prop() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Response.Cookies(\"demo\") = \"testval\"\nResponse.Cookies(\"demo\").Expires = \"2026-01-01\"",
                &mut ctx,
            )
            .unwrap();
        let entry = ctx.response.cookies.get("demo").unwrap();
        assert_eq!(entry.value.as_str(), "testval");
        assert_eq!(entry.expires, "2026-01-01");
    }

    #[test]
    fn test_asp_response_cookies_set_path() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Response.Cookies(\"x\") = \"y\"\nResponse.Cookies(\"x\").Path = \"/app\"",
                &mut ctx,
            )
            .unwrap();
        let entry = ctx.response.cookies.get("x").unwrap();
        assert_eq!(entry.value.as_str(), "y");
        assert_eq!(entry.path, "/app");
    }

    #[test]
    fn test_asp_response_cookies_indexed_get_returns_object() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Response.Cookies(\"pref\") = \"dark\"\nResponse.Cookies(\"pref\").Expires = \"2026-06-01\"",
                &mut ctx,
            )
            .unwrap();
        // Read back .Expires as an R-value
        interp
            .execute("Dim e\ne = Response.Cookies(\"pref\").Expires", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("e"),
            Some(&VBValue::String("2026-06-01".to_string()))
        );
    }

    #[test]
    fn test_asp_response_cookies_read_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute("Response.Cookies(\"greeting\") = \"hello\"", &mut ctx)
            .unwrap();
        interp
            .execute("Dim v\nv = Response.Cookies(\"greeting\").Value", &mut ctx)
            .unwrap();
        assert_eq!(
            ctx.get_variable("v"),
            Some(&VBValue::String("hello".to_string()))
        );
    }

    #[test]
    fn test_asp_response_cookies_secure_flag() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Response.Cookies(\"s\") = \"v\"\nResponse.Cookies(\"s\").Secure = True",
                &mut ctx,
            )
            .unwrap();
        let entry = ctx.response.cookies.get("s").unwrap();
        assert!(entry.secure);
    }

    #[test]
    fn test_asp_objects_injected_globally() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Dim rType, sType, svType, aType\nrType = TypeName(Request)\nsType = TypeName(Response)\nsvType = TypeName(Server)\naType = TypeName(Application)", &mut ctx).unwrap();
        assert_eq!(
            ctx.get_variable("rType"),
            Some(&VBValue::String("Request".to_string()))
        );
        assert_eq!(
            ctx.get_variable("sType"),
            Some(&VBValue::String("Response".to_string()))
        );
        assert_eq!(
            ctx.get_variable("svType"),
            Some(&VBValue::String("Server".to_string()))
        );
        assert_eq!(
            ctx.get_variable("aType"),
            Some(&VBValue::String("Application".to_string()))
        );
    }

    // ===== CLASS METHOD TESTS =====

    #[test]
    fn test_class_function_method_returns_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class TestClass\nPublic Function GetValue\n GetValue = 42\nEnd Function\nEnd Class\n\
                 Set obj = New TestClass\nresult = obj.GetValue()",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(42.0))
        );
    }

    #[test]
    fn test_class_method_with_params() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class TestClass\nPublic Function Add(a, b)\n Add = a + b\nEnd Function\nEnd Class\n\
                 Set obj = New TestClass\nresult = obj.Add(3, 4)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(7.0))
        );
    }

    #[test]
    fn test_class_method_mutates_instance_var() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class Counter\nPublic count\nPublic Sub Increment\n count = count + 1\nEnd Sub\nEnd Class\n\
                 Set c = New Counter\nc.count = 5\nc.Increment\nresult = c.count",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(6.0))
        );
    }

    #[test]
    fn test_class_method_accesses_global_object() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class TestClass\nPublic Function GetApp\n Set GetApp = Application\nEnd Function\nEnd Class\n\
                 Set obj = New TestClass\nSet app = obj.GetApp()\natype = TypeName(app)",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("atype"),
            Some(&VBValue::String("Application".to_string()))
        );
    }

    #[test]
    fn test_class_property_get_returns_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class C\nPublic val\nPublic Property Get Value\n Value = val\nEnd Property\n\
                 End Class\nSet c = New C\nc.val = 42\nresult = c.Value",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(42.0))
        );
    }

    #[test]
    fn test_class_property_let_sets_value() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class C\nPrivate val\nPublic Property Get Value\n Value = val\nEnd Property\n\
                 Public Property Let Value(v)\n val = v\nEnd Property\nEnd Class\n\
                 Set c = New C\nc.Value = 42\nresult = c.Value",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(42.0))
        );
    }

    #[test]
    fn test_class_property_case_insensitive() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Class C\nPrivate val\nPublic Property Get Value\n Value = val\nEnd Property\n\
                 Public Property Let Value(v)\n val = v\nEnd Property\nEnd Class\n\
                 Set c = New C\nc.VALUE = 42\nresult = c.VALUE",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(
            ctx.get_variable("result"),
            Some(&VBValue::Number(42.0))
        );
    }

    // ===== HTTP INTEGRATION TESTS =====


    async fn serve_and_get(asp_dir: &str, request: &str) -> String {
        let config = crate::asp::config::Config {
            host: "127.0.0.1".to_string(),
            port: 0,
            folder: asp_dir.to_string(),
            program: None,
            enable_directory_listing: false,
            default_documents: None,
            log_level: None,
        };
        let server = crate::asp::server::AspServer::new(config);
        let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();

        let folder = asp_dir.to_string();
        let store = Arc::clone(&server.store);
        let (shutdown_tx, mut shutdown_rx) = tokio::sync::oneshot::channel::<()>();

        let _handle = tokio::spawn(async move {
            loop {
                tokio::select! {
                    result = listener.accept() => {
                        let (stream, _) = result.unwrap();
                        let store = Arc::clone(&store);
                        let folder = folder.clone();
                        tokio::spawn(async move {
                            let mut stream = stream;
                            let dir_cache = crate::asp::config::DirConfigCache::new(
                                crate::asp::config::AspDirConfig {
                                    default_documents: vec!["index.asp".to_string()],
                                    directory_listing: false,
                                },
                                std::path::Path::new(&folder)
                                    .canonicalize()
                                    .unwrap_or_else(|_| std::path::Path::new(&folder).to_path_buf()),
                            );
                            let _ = crate::asp::server::AspServer::handle_connection(
                                &mut stream, &folder, &dir_cache, &store,
                            ).await;
                        });
                    }
                    _ = &mut shutdown_rx => break,
                }
            }
        });

        tokio::time::sleep(tokio::time::Duration::from_millis(50)).await;

        let mut client = tokio::net::TcpStream::connect(addr).await.unwrap();
        client.write_all(request.as_bytes()).await.unwrap();
        client.shutdown().await.unwrap();

        let mut reader = BufReader::new(&mut client);
        let mut response = String::new();
        let mut line = String::new();
        while reader.read_line(&mut line).await.unwrap() > 0 {
            response.push_str(&line);
            line.clear();
        }

        let _ = shutdown_tx.send(());
        tokio::time::sleep(tokio::time::Duration::from_millis(20)).await;
        response
    }

    #[tokio::test]
    async fn test_http_get_index_asp() {
        let dir = tmp_asp_dir();
        write_asp(
            &dir,
            "index.asp",
            "<html><body><%= \"Hello, World!\" %></body></html>",
        );

        let response = serve_and_get(
            dir.to_str().unwrap(),
            "GET / HTTP/1.1\r\nHost: localhost\r\n\r\n",
        )
        .await;

        assert!(
            response.contains("200 OK"),
            "Expected 200 OK, got: {}",
            response
        );
        assert!(
            response.contains("Hello, World!"),
            "Expected Hello, World!, got: {}",
            response
        );

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_404() {
        let dir = tmp_asp_dir();

        let response = serve_and_get(
            dir.to_str().unwrap(),
            "GET /nonexistent.asp HTTP/1.1\r\nHost: localhost\r\n\r\n",
        )
        .await;

        assert!(!response.is_empty(), "Empty response");
        assert!(response.contains("404"), "Expected 404, got: {}", response);

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_403() {
        let dir = tmp_asp_dir();
        write_asp(&dir, "index.asp", "OK");

        // Attempt path traversal (path won't canonicalize, so 404 first)
        let response = serve_and_get(
            dir.to_str().unwrap(),
            "GET /../Cargo.toml HTTP/1.1\r\nHost: localhost\r\n\r\n",
        )
        .await;

        assert!(!response.is_empty(), "Empty response");
        // Should get either 404 (canonicalize fails) or 403 (path traversal detected)
        assert!(
            response.contains("404") || response.contains("403"),
            "Expected 4xx, got: {}",
            response
        );

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_post_form() {
        let dir = tmp_asp_dir();
        write_asp(&dir, "index.asp", "<%= Request.Form(\"name\") %>");

        let body = "name=World";
        let request = format!(
            "POST / HTTP/1.1\r\nHost: localhost\r\nContent-Type: application/x-www-form-urlencoded\r\nContent-Length: {}\r\n\r\n{}",
            body.len(),
            body
        );
        let response = serve_and_get(dir.to_str().unwrap(), &request).await;

        assert!(
            response.contains("200 OK"),
            "Expected 200 OK, got: {}",
            response
        );
        let resp_body = response.split("\r\n\r\n").nth(1).unwrap_or("");
        assert!(
            resp_body.contains("World"),
            "Expected body to contain 'World', got: '{}'",
            resp_body
        );

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_multipart_form() {
        let dir = tmp_asp_dir();
        write_asp(&dir, "index.asp", "<%= Request.Form(\"field1\") %>");

        let boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW";
        let body = "------WebKitFormBoundary7MA4YWxkTrZu0gW\r\nContent-Disposition: form-data; name=\"field1\"\r\n\r\nvalue1\r\n------WebKitFormBoundary7MA4YWxkTrZu0gW--\r\n";
        let request = format!(
            "POST / HTTP/1.1\r\nHost: localhost\r\nContent-Type: multipart/form-data; boundary={}\r\nContent-Length: {}\r\n\r\n{}",
            boundary,
            body.len(),
            body
        );
        let response = serve_and_get(dir.to_str().unwrap(), &request).await;

        assert!(
            response.contains("200 OK"),
            "Expected 200 OK, got: {}",
            response
        );
        let resp_body = response.split("\r\n\r\n").nth(1).unwrap_or("");
        assert!(
            resp_body.contains("value1"),
            "Expected body to contain 'value1', got: '{}'",
            resp_body
        );

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_redirect() {
        let dir = tmp_asp_dir();
        write_asp(
            &dir,
            "redirect.asp",
            "<% Response.Redirect(\"other.asp\") %>",
        );

        let response = serve_and_get(
            dir.to_str().unwrap(),
            "GET /redirect.asp HTTP/1.1\r\nHost: localhost\r\n\r\n",
        )
        .await;

        assert!(response.contains("302"), "Expected 302, got: {}", response);
        assert!(
            response.contains("Location: other.asp"),
            "Expected Location header, got: {}",
            response
        );

        cleanup_dir(&dir);
    }

    #[tokio::test]
    async fn test_http_session_cookie() {
        let dir = tmp_asp_dir();
        write_asp(&dir, "index.asp", "Session.SessionID");

        let response = serve_and_get(
            dir.to_str().unwrap(),
            "GET / HTTP/1.1\r\nHost: localhost\r\n\r\n",
        )
        .await;

        assert!(
            response.contains("200 OK"),
            "Expected 200 OK, got: {}",
            response
        );
        assert!(
            response.contains("Set-Cookie: ASPSESSIONID=")
                || response.contains("Set-Cookie:ASPSESSIONID="),
            "Expected Set-Cookie header, got: {}",
            response
        );

        cleanup_dir(&dir);
    }
