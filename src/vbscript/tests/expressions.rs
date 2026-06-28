use super::*;
    // ===== EXPRESSION PARSER =====

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
                op: UnaryOp::Neg,
                expr: Box::new(Expr::Literal(VBValue::Number(5.0))),
            }
        );
    }

    #[test]
    fn test_parse_not_operator() {
        let tokens = Tokenizer::tokenize("Not True");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(
            expr,
            Expr::UnaryOp {
                op: UnaryOp::Not,
                expr: Box::new(Expr::Literal(VBValue::Boolean(true))),
            }
        );
    }

    #[test]
    fn test_parse_chained_comparison() {
        let tokens = Tokenizer::tokenize("1 < 2 And 3 > 1");
        let expr = parse_expression(&tokens).unwrap();
        assert_eq!(
            expr,
            Expr::BinaryOp {
                left: Box::new(Expr::BinaryOp {
                    left: Box::new(Expr::Literal(VBValue::Number(1.0))),
                    op: BinOp::Lt,
                    right: Box::new(Expr::Literal(VBValue::Number(2.0))),
                }),
                op: BinOp::And,
                right: Box::new(Expr::BinaryOp {
                    left: Box::new(Expr::Literal(VBValue::Number(3.0))),
                    op: BinOp::Gt,
                    right: Box::new(Expr::Literal(VBValue::Number(1.0))),
                }),
            }
        );
    }

    #[test]
    fn test_parse_lone_keyword_identifier() {
        // Keywords like "Null", "Empty", "True" should parse as literals
        let expr = parse_expression(&Tokenizer::tokenize("Null")).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::Null));

        let expr = parse_expression(&Tokenizer::tokenize("Empty")).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::Empty));

        let expr = parse_expression(&Tokenizer::tokenize("True")).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::Boolean(true)));

        let expr = parse_expression(&Tokenizer::tokenize("False")).unwrap();
        assert_eq!(expr, Expr::Literal(VBValue::Boolean(false)));
    }

    #[test]
    fn test_parse_unclosed_paren_error() {
        let tokens = Tokenizer::tokenize("(1 + 2");
        assert!(parse_expression(&tokens).is_err());
    }

    // ===== EXPRESSION EVALUATOR — ARITHMETIC =====

    #[test]
    fn test_evaluate_literal() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::Literal(VBValue::Number(42.0));
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Number(42.0)
        );
    }

    #[test]
    fn test_evaluate_variable() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("x", VBValue::Number(10.0));
        let expr = Expr::Variable("x".into());
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Number(10.0)
        );
    }

    #[test]
    fn test_evaluate_undefined_variable_error() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::Variable("undefined".into());
        assert!(evaluate(&expr, &mut context).is_err());
    }

    #[test]
    fn test_evaluate_add() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(3.0));
    }

    #[test]
    fn test_evaluate_sub() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Sub,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(7.0));
    }

    #[test]
    fn test_evaluate_mul() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(3.0))),
            op: BinOp::Mul,
            right: Box::new(Expr::Literal(VBValue::Number(4.0))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Number(12.0)
        );
    }

    #[test]
    fn test_evaluate_div() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Div,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_division_by_zero() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Div,
            right: Box::new(Expr::Literal(VBValue::Number(0.0))),
        };
        assert!(evaluate(&expr, &mut context).is_err());
    }

    #[test]
    fn test_evaluate_power() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Pow,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(8.0));
    }

    #[test]
    fn test_evaluate_int_divide() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::IntDiv,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(3.0));
    }

    #[test]
    fn test_evaluate_modulo() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Mod,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(1.0));
    }

    #[test]
    fn test_evaluate_unary_neg() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::UnaryOp {
            op: UnaryOp::Neg,
            expr: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Number(-5.0)
        );
    }

    #[test]
    fn test_evaluate_neg_of_negative() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::UnaryOp {
            op: UnaryOp::Neg,
            expr: Box::new(Expr::Literal(VBValue::Number(-3.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(3.0));
    }

    // ===== EXPRESSION EVALUATOR — STRING CONCAT =====

    #[test]
    fn test_evaluate_concat() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("Hello ".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::String("World".into()))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::String("Hello World".into())
        );
    }

    #[test]
    fn test_evaluate_concat_number_coercion() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // "a" & 42 → "a42"
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::Number(42.0))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::String("a42".into())
        );
    }

    // ===== EXPRESSION EVALUATOR — BOOLEAN / LOGICAL =====

    #[test]
    fn test_evaluate_and() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::And,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(false)
        );
    }

    #[test]
    fn test_evaluate_or() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Or,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
    }

    #[test]
    fn test_evaluate_not() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::UnaryOp {
            op: UnaryOp::Not,
            expr: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(false)
        );
    }

    #[test]
    fn test_evaluate_xor() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // true XOR true = false
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Xor,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(false)
        );

        // true XOR false = true
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Xor,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(
            evaluate(&expr2, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
    }

    #[test]
    fn test_evaluate_eqv() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // true Eqv true = true
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Eqv,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
        // true Eqv false = false
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Eqv,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(
            evaluate(&expr2, &mut context).unwrap(),
            VBValue::Boolean(false)
        );
    }

    #[test]
    fn test_evaluate_imp() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // true Imp false = false
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Imp,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(false)
        );
        // false Imp anything = true
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(false))),
            op: BinOp::Imp,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(
            evaluate(&expr2, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
    }

    // ===== EXPRESSION EVALUATOR — COMPARISON =====

    #[test]
    fn test_evaluate_comparison_eq() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Eq,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
    }

    #[test]
    fn test_evaluate_comparison_ne() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Ne,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(true)
        );

        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Ne,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(
            evaluate(&expr2, &mut context).unwrap(),
            VBValue::Boolean(false)
        );
    }

    #[test]
    fn test_evaluate_comparison_lt_gt() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let lt = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Lt,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&lt, &mut context).unwrap(), VBValue::Boolean(true));

        let gt = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Gt,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(evaluate(&gt, &mut context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_comparison_le_ge() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let le = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Le,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&le, &mut context).unwrap(), VBValue::Boolean(true));

        let ge = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(3.0))),
            op: BinOp::Ge,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&ge, &mut context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_is_operator() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // Is compares values — two nulls are equivalent
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Null)),
            op: BinOp::Is,
            right: Box::new(Expr::Literal(VBValue::Null)),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::Boolean(true)
        );
    }

    // ===== EXPRESSION EVALUATOR — TYPE COERCION =====

    #[test]
    fn test_evaluate_add_string_coercion() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // Number + String → string concat in VBScript-semantics
        // Our implementation: if both are numeric, add; else concat
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        // String + Number → concat as strings
        let result = evaluate(&expr, &mut context).unwrap();
        assert_eq!(result, VBValue::String("a1".into()));
    }

    #[test]
    fn test_evaluate_empty_as_number() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // Empty acts as 0 in numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Empty)),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_empty_as_string() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // Empty acts as "" in string context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("x".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::Empty)),
        };
        assert_eq!(
            evaluate(&expr, &mut context).unwrap(),
            VBValue::String("x".into())
        );
    }

    #[test]
    fn test_evaluate_empty() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::Literal(VBValue::Empty);
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Empty);
    }

    #[test]
    fn test_evaluate_null_as_number() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // Null acts as 0 in numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Null)),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_boolean_as_number() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        // True = -1 in VBScript numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(evaluate(&expr, &mut context).unwrap(), VBValue::Number(0.0));

        // False = 0
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(false))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(
            evaluate(&expr2, &mut context).unwrap(),
            VBValue::Number(1.0)
        );
    }

    // ===== ASSIGNMENT =====

    #[test]
    fn test_assignment_literal_number() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::Literal(VBValue::Number(42.0));
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(42.0)));
    }

    #[test]
    fn test_assignment_expression() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(3.0)));
    }

    #[test]
    fn test_assignment_variable_copy() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("a", VBValue::Number(10.0));
        let expr = Expr::Variable("a".into());
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(&VBValue::Number(10.0)));
    }

    #[test]
    fn test_assignment_string() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let assignment =
            Assignment::new("s".into(), Expr::Literal(VBValue::String("hello".into())));
        assignment.execute(&mut context).unwrap();
        assert_eq!(
            context.get_variable("s"),
            Some(&VBValue::String("hello".into()))
        );
    }

    #[test]
    fn test_assignment_boolean() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let assignment = Assignment::new("b".into(), Expr::Literal(VBValue::Boolean(true)));
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("b"), Some(&VBValue::Boolean(true)));
    }

    #[test]
    fn test_assignment_null() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        let assignment = Assignment::new("n".into(), Expr::Literal(VBValue::Null));
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("n"), Some(&VBValue::Null));
    }

    #[test]
    fn test_assignment_concat_expression() {
        let mut context = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut context);
        context.set_variable("pre", VBValue::String("Hello ".into()));
        let assignment = Assignment::new(
            "msg".into(),
            Expr::BinaryOp {
                left: Box::new(Expr::Variable("pre".into())),
                op: BinOp::Concat,
                right: Box::new(Expr::Literal(VBValue::String("World".into()))),
            },
        );
        assignment.execute(&mut context).unwrap();
        assert_eq!(
            context.get_variable("msg"),
            Some(&VBValue::String("Hello World".into()))
        );
    }

