#[cfg(test)]
mod tests {
    use crate::asp::parser::AspParser;
    use crate::vbscript::expr::{evaluate, parse_expression, BinOp, Expr};
    use crate::vbscript::syntax::{Assignment, Dim, ResponseWrite, VBSyntax};
    use crate::vbscript::{ExecutionContext, Tokenizer, VBValue};

    // --- Expression Parser Tests ---

    #[test]
    fn test_parse_literal_number() {
        let tokens = Tokenizer::tokenize("42");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::Number(42.0)));
    }

    #[test]
    fn test_parse_literal_string() {
        let tokens = Tokenizer::tokenize(r#""hello""#);
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::String("hello".into())));
    }

    #[test]
    fn test_parse_variable() {
        let tokens = Tokenizer::tokenize("x");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(expr, Expr::Variable("x".into()));
    }

    #[test]
    fn test_parse_binary_add() {
        let tokens = Tokenizer::tokenize("1 + 2");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(
            expr,
            Expr::BinaryOp {
                left: Box::new(Expr::Literal(VBValue::Number(1.0))),
                op: BinOp::Add,
                right: Box::new(Expr::Literal(VBValue::Number(2.0))),
            }
        );
    }

    #[test]
    fn test_parse_precedence_mul_over_add() {
        let tokens = Tokenizer::tokenize("1 + 2 * 3");
        let expr = parse_expression(&tokens).unwrap();
        // Should be: 1 + (2 * 3)
        assert_eq!(
            expr,
            Expr::BinaryOp {
                left: Box::new(Expr::Literal(VBValue::Number(1.0))),
                op: BinOp::Add,
                right: Box::new(Expr::BinaryOp {
                    left: Box::new(Expr::Literal(VBValue::Number(2.0))),
                    op: BinOp::Mul,
                    right: Box::new(Expr::Literal(VBValue::Number(3.0))),
                }),
            }
        );
    }

    #[test]
    fn test_parse_parentheses() {
        let tokens = Tokenizer::tokenize("(1 + 2) * 3");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(
            expr,
            Expr::BinaryOp {
                left: Box::new(Expr::BinaryOp {
                    left: Box::new(Expr::Literal(VBValue::Number(1.0))),
                    op: BinOp::Add,
                    right: Box::new(Expr::Literal(VBValue::Number(2.0))),
                }),
                op: BinOp::Mul,
                right: Box::new(Expr::Literal(VBValue::Number(3.0))),
            }
        );
    }

    #[test]
    fn test_parse_unary_minus() {
        let tokens = Tokenizer::tokenize("-5");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(
            expr,
            Expr::UnaryOp {
                op: crate::vbscript::expr::UnaryOp::Neg,
                expr: Box::new(Expr::Literal(VBValue::Number(5.0))),
            }
        );
    }

    // --- Expression Evaluator Tests ---

    #[test]
    fn test_evaluate_literal() {
        let context = ExecutionContext::new();
        let expr = Expr::Literal(VBValue::Number(42.0));
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(42.0));
    }

    #[test]
    fn test_evaluate_variable() {
        let mut context = ExecutionContext::new();
        context.set_variable("x", VBValue::Number(10.0));
        let expr = Expr::Variable("x".into());
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(10.0));
    }

    #[test]
    fn test_evaluate_add() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(3.0));
    }

    #[test]
    fn test_evaluate_mul() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(3.0))),
            op: BinOp::Mul,
            right: Box::new(Expr::Literal(VBValue::Number(4.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(12.0));
    }

    #[test]
    fn test_evaluate_sub() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Sub,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(7.0));
    }

    #[test]
    fn test_evaluate_div() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Div,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_division_by_zero() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Div,
            right: Box::new(Expr::Literal(VBValue::Number(0.0))),
        };
        assert!(evaluate(&expr, &context).is_err());
    }

    #[test]
    fn test_evaluate_and() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::And,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_or() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Or,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_not() {
        let context = ExecutionContext::new();
        let expr = Expr::UnaryOp {
            op: crate::vbscript::expr::UnaryOp::Not,
            expr: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_concat() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("Hello ".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::String("World".into()))),
        };
        assert_eq!(
            evaluate(&expr, &context).unwrap(),
            VBValue::String("Hello World".into())
        );
    }

    #[test]
    fn test_evaluate_comparison_eq() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Eq,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_comparison_gt() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Gt,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_unary_neg() {
        let context = ExecutionContext::new();
        let expr = Expr::UnaryOp {
            op: crate::vbscript::expr::UnaryOp::Neg,
            expr: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(-5.0));
    }

    #[test]
    fn test_evaluate_empty() {
        let context = ExecutionContext::new();
        let expr = Expr::Literal(VBValue::Empty);
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Empty);
    }

    // --- Assignment Tests ---

    #[test]
    fn test_assignment_literal_number() {
        let mut context = ExecutionContext::new();
        let expr = Expr::Literal(VBValue::Number(42.0));
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(42.0)));
    }

    #[test]
    fn test_assignment_expression() {
        let mut context = ExecutionContext::new();
        // x = 1 + 2 → x should be 3
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(3.0)));
    }

    #[test]
    fn test_assignment_variable_copy() {
        let mut context = ExecutionContext::new();
        context.set_variable("a", VBValue::Number(10.0));
        // x = a
        let expr = Expr::Variable("a".into());
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(10.0)));
    }

    // --- ResponseWrite Tests ---

    #[test]
    fn test_response_write_literal() {
        let mut context = ExecutionContext::new();
        let expr = Expr::Literal(VBValue::String("hello".into()));
        let rw = ResponseWrite::new(expr);
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "hello");
    }

    #[test]
    fn test_response_write_variable() {
        let mut context = ExecutionContext::new();
        context.set_variable("name", VBValue::String("World".into()));
        let expr = Expr::Variable("name".into());
        let rw = ResponseWrite::new(expr);
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "World");
    }

    // --- Dim Tests ---

    #[test]
    fn test_dim_initializes_to_empty() {
        let mut context = ExecutionContext::new();
        let dim = Dim::new(vec!["x".into()]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Empty));
    }

    // --- Integration Tests ---

    #[test]
    fn test_integration_assignment_evaluation() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;

        // x = 5 * (2 + 1)
        interpreter.execute("x = 5 * (2 + 1)", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(15.0)));
    }

    #[test]
    fn test_integration_multiple_statements() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;

        interpreter.execute(r#"
            Dim name
            name = "test"
            x = 10 + 20
        "#, &mut context).unwrap();

        assert_eq!(context.get_variable("name"), Some(VBValue::String("test".into())));
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(30.0)));
    }

    #[test]
    fn test_integration_response_write() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;

        interpreter.execute(r#"Response.Write "hello""#, &mut context).unwrap();
        assert_eq!(context.response_buffer, "hello");
    }

    #[test]
    fn test_integration_variable_in_expression() {
        let mut context = ExecutionContext::new();
        context.set_variable("a", VBValue::Number(5.0));
        let interpreter = crate::vbscript::VBScriptInterpreter;

        interpreter.execute("x = a * 2", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(10.0)));
    }

    #[test]
    fn test_integration_concat_in_response_write() {
        let mut context = ExecutionContext::new();
        context.set_variable("name", VBValue::String("World".into()));
        let interpreter = crate::vbscript::VBScriptInterpreter;

        interpreter.execute(r#"Response.Write "Hello " & name"#, &mut context).unwrap();
        assert_eq!(context.response_buffer, "Hello World");
    }

    #[test]
    fn test_asp_parser_splits_html_and_code() {
        let parser = AspParser::new("<html><%Dim x%></html>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 3); // html, code, html
        match &blocks[0] {
            crate::asp::parser::AspBlock::Html(h) => assert_eq!(h, "<html>"),
            _ => panic!("Expected Html block"),
        }
        match &blocks[1] {
            crate::asp::parser::AspBlock::Code(c) => assert_eq!(c, "Dim x"),
            _ => panic!("Expected Code block"),
        }
    }

    // --- Block Statement Tests ---

    #[test]
    fn test_if_inline_true() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("If x = 1 Then y = 42", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(42.0)));
    }

    #[test]
    fn test_if_inline_false() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(2.0));
        interpreter.execute("If x = 1 Then y = 42", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), None);
    }

    #[test]
    fn test_if_block_true() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("If x = 1 Then\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(99.0)));
    }

    #[test]
    fn test_if_block_false() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter.execute("If x = 1 Then\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), None);
    }

    #[test]
    fn test_if_else() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter.execute("If x = 1 Then\n    y = 10\nElse\n    y = 20\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(20.0)));
    }

    #[test]
    fn test_for_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim total\ntotal = 0\nFor i = 1 To 5\n    total = total + i\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(15.0)));
    }

    #[test]
    fn test_while_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("While x <= 3\n    x = x + 1\nWend", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(4.0)));
    }

    #[test]
    fn test_do_while() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("Do While x < 3\n    x = x + 1\nLoop", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(3.0)));
    }

    #[test]
    fn test_do_loop_until() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("Do\n    x = x + 1\nLoop Until x > 3", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(4.0)));
    }

    #[test]
    fn test_for_with_step() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim total\ntotal = 0\nFor i = 1 To 10 Step 2\n    total = total + i\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(25.0)));
    }

    #[test]
    fn test_nested_blocks() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim result\nresult = 0\nFor i = 1 To 3\n    If i > 1 Then\n        result = result + i\n    End If\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("result"), Some(VBValue::Number(5.0)));
    }

    #[test]
    fn test_response_write_preserved() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Response.Write \"hello\"", &mut context).unwrap();
        assert_eq!(context.response_buffer, "hello");
    }

    #[test]
    fn test_dim_preserved() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim x", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Empty));
    }
}
