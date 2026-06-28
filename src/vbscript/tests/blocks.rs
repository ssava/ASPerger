use super::*;
    // ===== RESPONSE.WRITE =====

    #[test]
    fn test_response_write_literal() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let rw = ResponseWrite::new(Expr::Literal(VBValue::String("hello".into())));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response.buffer, "hello");
    }

    #[test]
    fn test_response_write_variable() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("name", VBValue::String("World".into()));
        let rw = ResponseWrite::new(Expr::Variable("name".into()));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response.buffer, "World");
    }

    #[test]
    fn test_response_write_number() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let rw = ResponseWrite::new(Expr::Literal(VBValue::Number(42.0)));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response.buffer, "42");
    }

    #[test]
    fn test_response_write_expression() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let rw = ResponseWrite::new(Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::String("b".into()))),
        });
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response.buffer, "ab");
    }

    // ===== DIM =====

    #[test]
    fn test_dim_initializes_to_empty() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let dim = Dim::new(vec![("x".into(), None)]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Empty));
    }

    #[test]
    fn test_dim_multiple_variables() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let dim = Dim::new(vec![
            ("a".into(), None),
            ("b".into(), None),
            ("c".into(), None),
        ]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("a"), Some(&VBValue::Empty));
        assert_eq!(context.get_variable("b"), Some(&VBValue::Empty));
        assert_eq!(context.get_variable("c"), Some(&VBValue::Empty));
    }

    // ===== REDIM =====

    #[test]
    fn test_redim_basic() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Dim arr\nReDim arr(4)\narr(0) = \"a\"\narr(4) = \"e\"", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("arr"), Some(&VBValue::Array(
            std::sync::Arc::new(vec![
                VBValue::String("a".into()),
                VBValue::Empty,
                VBValue::Empty,
                VBValue::Empty,
                VBValue::String("e".into()),
            ]),
            vec![4],
        )));
    }

    #[test]
    fn test_redim_multi_dim() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Dim grid\nReDim grid(2, 3)", &mut context)
            .unwrap();
        // 3 rows (0..2) × 4 cols (0..3) = 12 elements
        assert_eq!(context.get_variable("grid"), Some(&VBValue::Array(
            std::sync::Arc::new(vec![VBValue::Empty; 12]),
            vec![2, 3],
        )));
    }

    #[test]
    fn test_redim_multi_dim_access() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Dim grid\nReDim grid(2, 3)\ngrid(1, 2) = \"hello\"", &mut context)
            .unwrap();
        // flat index: 1 * (3+1) + 2 = 6
        let expected = VBValue::Array(
            std::sync::Arc::new({
                let mut v = vec![VBValue::Empty; 12];
                v[6] = VBValue::String("hello".into());
                v
            }),
            vec![2, 3],
        );
        assert_eq!(context.get_variable("grid"), Some(&expected));
    }

    #[test]
    fn test_redim_multi_dim_expression_access() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Dim grid\nReDim grid(2, 3)\ngrid(1, 2) = \"hello\"\nresult = grid(1, 2)", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("result"), Some(&VBValue::String("hello".into())));
    }

    #[test]
    fn test_redim_preserve_multi_dim_error() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("Dim arr\nReDim arr(2, 3)\nReDim Preserve arr(4, 5)", &mut context);
        assert!(result.is_err());
    }

    // ===== BLOCK STATEMENTS — IF =====

    #[test]
    fn test_if_inline_true() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("If x = 1 Then y = 42", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(42.0)));
    }

    #[test]
    fn test_if_inline_false() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(2.0));
        interpreter
            .execute("If x = 1 Then y = 42", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), None);
    }

    #[test]
    fn test_if_block_true() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("If x = 1 Then\n    y = 99\nEnd If", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(99.0)));
    }

    #[test]
    fn test_if_block_false() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter
            .execute("If x = 1 Then\n    y = 99\nEnd If", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), None);
    }

    #[test]
    fn test_if_else() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter
            .execute(
                "If x = 1 Then\n    y = 10\nElse\n    y = 20\nEnd If",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(20.0)));
    }

    #[test]
    fn test_if_elseif_chain() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // First ElseIf matches
        context.set_variable("x", VBValue::Number(2.0));
        interpreter.execute("If x = 1 Then\n    y = 10\nElseIf x = 2 Then\n    y = 20\nElseIf x = 3 Then\n    y = 30\nElse\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(20.0)));
    }

    #[test]
    fn test_if_elseif_falls_to_else() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(9.0));
        interpreter.execute("If x = 1 Then\n    y = 10\nElseIf x = 2 Then\n    y = 20\nElse\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(99.0)));
    }

    #[test]
    fn test_if_multiple_blocks_in_sequence() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "If 1 = 1 Then\n    a = 10\nEnd If\nIf 2 = 2 Then\n    b = 20\nEnd If",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("a"), Some(&VBValue::Number(10.0)));
        assert_eq!(context.get_variable("b"), Some(&VBValue::Number(20.0)));
    }

    #[test]
    fn test_if_condition_with_and() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(5.0));
        interpreter
            .execute("If x > 1 And x < 10 Then y = 1\nEnd If", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(1.0)));
    }

    // ===== BLOCK STATEMENTS — FOR =====

    #[test]
    fn test_for_loop() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Dim total\ntotal = 0\nFor i = 1 To 5\n    total = total + i\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_for_with_step() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Dim total\ntotal = 0\nFor i = 1 To 10 Step 2\n    total = total + i\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(25.0)));
    }

    #[test]
    fn test_for_negative_step() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute(
                "Dim total\ntotal = 0\nFor i = 5 To 1 Step -1\n    total = total + i\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(15.0)));
    }

    #[test]
    fn test_for_zero_iterations() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // start > end with positive step → zero iterations
        interpreter
            .execute(
                "Dim total\ntotal = 99\nFor i = 5 To 1\n    total = 0\nNext",
                &mut context,
            )
            .unwrap();
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(99.0)));
    }

    #[test]
    fn test_for_empty_body() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("For i = 1 To 3\nNext", &mut context)
            .unwrap();
        // Just shouldn't crash; counter should be 4 (past end)
        assert_eq!(context.get_variable("i"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_for_modifies_counter_in_body() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // Modifying the loop counter inside the body
        interpreter
            .execute(
                "Dim total\ntotal = 0\nFor i = 1 To 10\n    total = total + i\n    i = i + 1\nNext",
                &mut context,
            )
            .unwrap();
        // i increments by 1 in loop header, plus +1 in body = effective step of 2
        // So i goes: 1 → (header: 2, body: +1) → 3 → 4 → 5 → 6 → 7 → 8 → 9 → header:10 → body:11 → header:12 stops
        // The loop counter is managed by the For statement,
        // so modifying i in the body doesn't affect iteration count.
        assert_eq!(context.get_variable("total"), Some(&VBValue::Number(55.0)));
    }

    // ===== BLOCK STATEMENTS — WHILE =====

    #[test]
    fn test_while_loop() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("While x <= 3\n    x = x + 1\nWend", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_while_never_enters() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter
            .execute("While x < 5\n    x = 99\nWend", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(10.0)));
    }

    #[test]
    fn test_while_empty_body_does_not_loop() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter
            .execute("While x > 0\n    x = 99\nWend", &mut context)
            .unwrap();
        // Should not modify x since condition is false
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(0.0)));
    }

    // ===== BLOCK STATEMENTS — DO =====

    #[test]
    fn test_do_while() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("Do While x < 3\n    x = x + 1\nLoop", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(3.0)));
    }

    #[test]
    fn test_do_while_never_enters() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter
            .execute("Do While x < 5\n    x = 99\nLoop", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(10.0)));
    }

    #[test]
    fn test_do_loop_until() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("Do\n    x = x + 1\nLoop Until x > 3", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_do_loop_while_post_test() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter
            .execute("Do\n    x = x + 1\nLoop While x < 3", &mut context)
            .unwrap();
        // Post-test: executes body, then checks. So: x=2 (check 2<3), x=3 (check 3<3=false)
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(3.0)));
    }

    #[test]
    fn test_do_loop_until_post_test_runs_once() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter
            .execute("Do\n    x = x + 1\nLoop Until x > 5", &mut context)
            .unwrap();
        // Post-test: always runs at least once. x=11 (check 11>5=true)
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(11.0)));
    }

    // ===== NESTED BLOCKS =====

    #[test]
    fn test_nested_blocks() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim result\nresult = 0\nFor i = 1 To 3\n    If i > 1 Then\n        result = result + i\n    End If\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("result"), Some(&VBValue::Number(5.0)));
    }

    #[test]
    fn test_deeply_nested_blocks() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute(
            "Dim out\nout = 0\nFor a = 1 To 2\n    For b = 1 To 2\n        If b = a Then\n            out = out + 1\n        End If\n    Next\nNext",
            &mut context,
        )
        .unwrap();
        // a=1,b=1: matches → out=1; a=1,b=2: no; a=2,b=1: no; a=2,b=2: matches → out=2
        assert_eq!(context.get_variable("out"), Some(&VBValue::Number(2.0)));
    }

    // ===== COMMENTS =====

    #[test]
    fn test_comment_apostrophe_line() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("' this is a comment\nx = 1", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_comment_rem_keyword() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Rem this is a comment\ny = 2", &mut context)
            .unwrap();
        assert_eq!(context.get_variable("y"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_code_with_only_comments() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("' comment 1\n' comment 2\nRem comment 3", &mut context)
            .unwrap();
        // Just shouldn't crash
    }

    // ===== ERROR HANDLING =====

    #[test]
    fn test_undefined_variable_in_expression() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("x = undefinedVar + 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_division_by_zero_in_if_condition() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("If 1 / 0 = 1 Then x = 1\nEnd If", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_if_without_then() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("If x = 1\nEnd If", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_for_without_next() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("For i = 1 To 5\nx = 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_while_without_wend() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("While x < 5\nx = x + 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_do_without_loop() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("Do While x < 5\nx = x + 1", &mut context);
        assert!(result.is_err());
    }

    // ===== EMPTY / WHITESPACE CODE =====

    #[test]
    fn test_empty_code() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("", &mut context).unwrap();
        // No crash
    }

    #[test]
    fn test_whitespace_code() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("   \n  \t  \n  ", &mut context)
            .unwrap();
        // No crash
    }

    // ===== PRESERVED BEHAVIOR =====

    #[test]
    fn test_response_write_preserved() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter
            .execute("Response.Write \"hello\"", &mut context)
            .unwrap();
        assert_eq!(context.response.buffer, "hello");
    }

    #[test]
    fn test_dim_preserved() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim x", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Empty));
    }

    // ===== EXIT STATEMENTS =====

    #[test]
    fn test_exit_for() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute("Dim i, sum\nsum = 0\nFor i = 1 To 10\n    If i = 5 Then Exit For\n    sum = sum + 1\nNext", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("sum"), Some(&VBValue::Number(4.0)));
        assert_eq!(ctx.get_variable("i"), Some(&VBValue::Number(5.0)));
    }

    #[test]
    fn test_exit_for_from_foreach() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_variable(
            "items",
            VBValue::Array(std::sync::Arc::new(vec![
                VBValue::Number(1.0),
                VBValue::Number(2.0),
                VBValue::Number(3.0),
                VBValue::Number(4.0),
                VBValue::Number(5.0),
            ]), vec![]),
        );
        interp.execute("Dim x, sum\nsum = 0\nFor Each x In items\n    If x = 3 Then Exit For\n    sum = sum + 1\nNext", &mut ctx).unwrap();
        assert_eq!(ctx.get_variable("sum"), Some(&VBValue::Number(2.0)));
    }

    #[test]
    fn test_exit_do_while() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_variable("x", VBValue::Number(1.0));
        interp
            .execute(
                "Do While x < 10\n    If x = 5 Then Exit Do\n    x = x + 1\nLoop",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(5.0)));
    }

    #[test]
    fn test_exit_do_loop_until() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_variable("x", VBValue::Number(1.0));
        interp
            .execute(
                "Do\n    If x = 3 Then Exit Do\n    x = x + 1\nLoop While x < 10",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(3.0)));
    }

    #[test]
    fn test_exit_for_nested_only_exits_inner() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute(
            "Dim i, j, sum\nsum = 0\nFor i = 1 To 3\n    For j = 1 To 3\n        If j = 2 Then Exit For\n        sum = sum + 1\n    Next\nNext",
            &mut ctx
        ).unwrap();
        // inner loop: j=1 then exits at j=2 for each i
        // i=1: j=1 -> sum=1, exit j=2
        // i=2: j=1 -> sum=2, exit j=2
        // i=3: j=1 -> sum=3, exit j=2
        assert_eq!(ctx.get_variable("sum"), Some(&VBValue::Number(3.0)));
        // outer loop completed all 3 iterations
        assert_eq!(ctx.get_variable("i"), Some(&VBValue::Number(4.0)));
    }

    #[test]
    fn test_exit_function_early_return() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute(
            "Function TestFunc(x)\n    If x > 5 Then\n        TestFunc = 999\n        Exit Function\n    End If\n    TestFunc = 111\nEnd Function\nresult = TestFunc(10)",
            &mut ctx
        ).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(999.0)));
    }

    #[test]
    fn test_exit_function_normal_path() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp.execute(
            "Function TestFunc(x)\n    If x > 5 Then\n        TestFunc = 999\n        Exit Function\n    End If\n    TestFunc = 111\nEnd Function\nresult = TestFunc(3)",
            &mut ctx
        ).unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(111.0)));
    }

    #[test]
    fn test_exit_sub() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        ctx.set_variable("called", VBValue::Boolean(false));
        interp
            .execute(
                "Sub TestSub()\n    Exit Sub\n    called = True\nEnd Sub\nCall TestSub()",
                &mut ctx,
            )
            .unwrap();
        // called should still be false since Exit Sub prevented the assignment
        assert_eq!(ctx.get_variable("called"), Some(&VBValue::Boolean(false)));
    }

    // ===== WITH / END WITH =====

    #[test]
    fn test_with_basic() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute(
                "With dict\n    .Add \"key\", \"Alpha\"\n    result = .Count\nEnd With",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_with_property_access() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut ctx,
            )
            .unwrap();
        interp.execute("dict.Add \"a\", \"1\"", &mut ctx).unwrap();
        interp
            .execute("With dict\n    result = .Count\nEnd With", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(1.0)));
    }

    #[test]
    fn test_with_method_call() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut ctx,
            )
            .unwrap();
        interp.execute("dict.Add \"a\", \"1\"", &mut ctx).unwrap();
        interp.execute("dict.Add \"b\", \"2\"", &mut ctx).unwrap();
        interp
            .execute(
                "With dict\n    .RemoveAll\n    result = .Count\nEnd With",
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(0.0)));
    }

    #[test]
    fn test_with_property_get() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set dict = CreateObject(\"Scripting.Dictionary\")",
                &mut ctx,
            )
            .unwrap();
        interp
            .execute("dict.Add \"a\", \"Alpha\"", &mut ctx)
            .unwrap();
        interp
            .execute("dict.Add \"b\", \"Beta\"", &mut ctx)
            .unwrap();
        interp
            .execute("With dict\n    result = .Count\nEnd With", &mut ctx)
            .unwrap();
        assert_eq!(ctx.get_variable("result"), Some(&VBValue::Number(2.0)));
    }
