#[cfg(test)]
mod tests {
    use crate::asp::parser::AspParser;
    use crate::vbscript::expr::{evaluate, parse_expression, BinOp, Expr, UnaryOp};
    use crate::vbscript::syntax::{Assignment, Dim, ResponseWrite, VBSyntax};
    use crate::vbscript::{ExecutionContext, TokenType, Tokenizer, VBScriptInterpreter, VBValue};
    use chrono::{Datelike, Timelike};
    use std::fs;
    use std::sync::Arc;

    // ===== TOKENIZER =====

    #[test]
    fn test_tokenizer_simple_assignment() {
        let tokens = Tokenizer::tokenize("x = 5");
        assert_eq!(tokens.len(), 4);
        assert_eq!(tokens[0].token_type, TokenType::Identifier);
        assert_eq!(tokens[0].value.as_ref(), "x");
        assert_eq!(tokens[1].token_type, TokenType::Assign);
        assert_eq!(tokens[2].token_type, TokenType::IntegerLiteral);
        assert_eq!(tokens[2].value.as_ref(), "5");
        assert_eq!(tokens[3].token_type, TokenType::EOF);
    }

    #[test]
    fn test_tokenizer_keywords() {
        let tokens = Tokenizer::tokenize("If Then Else ElseIf End For Next Do Loop While Wend");
        let expected = [
            TokenType::If,
            TokenType::Then,
            TokenType::Else,
            TokenType::ElseIf,
            TokenType::End,
            TokenType::For,
            TokenType::Next,
            TokenType::Do,
            TokenType::Loop,
            TokenType::While,
            TokenType::WEnd,
        ];
        for (i, expected_tt) in expected.iter().enumerate() {
            assert_eq!(
                tokens[i].token_type, *expected_tt,
                "mismatch at token {}: got {:?}",
                i, tokens[i].token_type
            );
        }
    }

    #[test]
    fn test_tokenizer_operator_keywords() {
        let tokens = Tokenizer::tokenize("And Or Not Mod Is Eqv Imp To Step");
        let expected = [
            TokenType::And,
            TokenType::Or,
            TokenType::Not,
            TokenType::Mod,
            TokenType::Is,
            TokenType::Eqv,
            TokenType::Imp,
            TokenType::To,
            TokenType::Step,
        ];
        for (i, expected_tt) in expected.iter().enumerate() {
            assert_eq!(
                tokens[i].token_type, *expected_tt,
                "mismatch at token {}: got {:?}",
                i, tokens[i].token_type
            );
        }
    }

    #[test]
    fn test_tokenizer_string_with_escaped_quotes() {
        let tokens = Tokenizer::tokenize(r#""he""llo""#);
        assert_eq!(tokens[0].token_type, TokenType::StringLiteral);
        assert_eq!(tokens[0].value.as_ref(), r#"he"llo"#);
    }

    #[test]
    fn test_tokenizer_comment() {
        let tokens = Tokenizer::tokenize("' this is a comment");
        assert_eq!(tokens[0].token_type, TokenType::Comment);
        assert_eq!(tokens[0].value.as_ref(), " this is a comment");
        assert_eq!(tokens[1].token_type, TokenType::EOF);
    }

    #[test]
    fn test_tokenizer_hex_literal() {
        let tokens = Tokenizer::tokenize("&HFF");
        assert_eq!(tokens[0].token_type, TokenType::HexLiteral);
        assert_eq!(tokens[0].value.as_ref(), "&HFF");
    }

    #[test]
    fn test_tokenizer_oct_literal() {
        let tokens = Tokenizer::tokenize("&77");
        assert_eq!(tokens[0].token_type, TokenType::OctLiteral);
        assert_eq!(tokens[0].value.as_ref(), "&77");
    }

    #[test]
    fn test_tokenizer_float() {
        let tokens = Tokenizer::tokenize("3.14");
        assert_eq!(tokens[0].token_type, TokenType::FloatLiteral);
        assert_eq!(tokens[0].value.as_ref(), "3.14");
    }

    #[test]
    fn test_tokenizer_scientific_notation() {
        let tokens = Tokenizer::tokenize("1.5e2");
        assert_eq!(tokens[0].token_type, TokenType::FloatLiteral);
        assert_eq!(tokens[0].value.as_ref(), "1.5e2");
    }

    #[test]
    fn test_tokenizer_date_literal() {
        let tokens = Tokenizer::tokenize("#2024-01-15#");
        assert_eq!(tokens[0].token_type, TokenType::DateLiteral);
        assert_eq!(tokens[0].value.as_ref(), "2024-01-15");
    }

    #[test]
    fn test_tokenizer_comparison_equal_vs_assign() {
        let tokens = Tokenizer::tokenize("x = 5");
        assert_eq!(tokens[1].token_type, TokenType::Assign);
        let tokens2 = Tokenizer::tokenize("x == 5");
        assert_eq!(tokens2[1].token_type, TokenType::Equal);
        assert_eq!(tokens2[1].value.as_ref(), "==");
    }

    #[test]
    fn test_tokenizer_operators() {
        let tokens = Tokenizer::tokenize("+ - * / \\ ^ & ( ) , . : > < >= <= <>");
        let expected = [
            TokenType::Plus,
            TokenType::Minus,
            TokenType::Multiply,
            TokenType::Divide,
            TokenType::IntDivide,
            TokenType::Power,
            TokenType::Concat,
            TokenType::LeftParen,
            TokenType::RightParen,
            TokenType::Comma,
            TokenType::Dot,
            TokenType::Colon,
            TokenType::GreaterThan,
            TokenType::LessThan,
            TokenType::GreaterEqual,
            TokenType::LessEqual,
            TokenType::NotEqual,
        ];
        for (i, expected_tt) in expected.iter().enumerate() {
            assert_eq!(
                tokens[i].token_type, *expected_tt,
                "mismatch at token {}: got {:?} value={:?}",
                i, tokens[i].token_type, tokens[i].value
            );
        }
    }

    #[test]
    fn test_tokenizer_empty_input() {
        let tokens = Tokenizer::tokenize("");
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0].token_type, TokenType::EOF);
    }

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
            Some(&VBValue::String("Alpha".to_string()))
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
        context.set_variable("result", VBValue::String("".to_string()));
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
            VBValue::String(s) => s.clone(),
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

        for block in &blocks {
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

    // ===== FILESYSTEMOBJECT + TEXTSTREAM =====

    fn tmp_path(name: &str) -> String {
        let p = std::env::temp_dir().join(format!("asperger_test_{}", name));
        p.to_str().unwrap().to_string()
    }

    fn cleanup_path(path: &str) {
        let p = std::path::Path::new(path);
        if p.is_file() {
            let _ = std::fs::remove_file(p);
        } else if p.is_dir() {
            let _ = std::fs::remove_dir_all(p);
        }
    }

    #[test]
    fn test_fso_createobject() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")",
                &mut ctx,
            )
            .unwrap();
        let fso = ctx.get_variable("fso");
        assert!(fso.is_some());
        assert!(matches!(fso.unwrap(), VBValue::Object(_)));
    }

    #[test]
    fn test_fso_fileexists() {
        let path = tmp_path("fileexists.txt");
        cleanup_path(&path);
        assert!(!std::path::Path::new(&path).exists());

        // Create the file
        std::fs::File::create(&path).unwrap();
        assert!(std::path::Path::new(&path).exists());

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FileExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(true)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folderexists() {
        let path = tmp_path("folderexists");
        cleanup_path(&path);
        std::fs::create_dir_all(&path).unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FolderExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(true)));

        // Non-existent folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\nx = fso.FolderExists(\"{}_nonexistent\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("x"), Some(&VBValue::Boolean(false)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_createtextfile_and_readall() {
        let path = tmp_path("create_read.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Create, write, close
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.Write \"Hello, World!\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // Open and read all
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     content = ts.ReadAll()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("content"),
            Some(&VBValue::String("Hello, World!".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_writeline_and_readline() {
        let path = tmp_path("writeline.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Write multiple lines
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.WriteLine \"Line 1\"\n\
                     ts.WriteLine \"Line 2\"\n\
                     ts.WriteLine \"Line 3\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // Read them back
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     line1 = ts.ReadLine()\n\
                     line2 = ts.ReadLine()\n\
                     line3 = ts.ReadLine()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("line1"),
            Some(&VBValue::String("Line 1".to_string()))
        );
        assert_eq!(
            ctx.get_variable("line2"),
            Some(&VBValue::String("Line 2".to_string()))
        );
        assert_eq!(
            ctx.get_variable("line3"),
            Some(&VBValue::String("Line 3".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_read_n_characters() {
        let path = tmp_path("readchars.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        std::fs::write(&path, "ABCDEFGHIJ").unwrap();

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     part = ts.Read(4)\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("part"),
            Some(&VBValue::String("ABCD".to_string()))
        );

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_atendofstream() {
        let path = tmp_path("atend.txt");
        cleanup_path(&path);

        std::fs::write(&path, "Hello").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 1)\n\
                     ' Should not be at end initially\n\
                     initial = ts.AtEndOfStream\n\
                     content = ts.ReadAll()\n\
                     ' Should be at end after reading all\n\
                     after = ts.AtEndOfStream\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("initial"), Some(&VBValue::Boolean(false)));
        assert_eq!(ctx.get_variable("after"), Some(&VBValue::Boolean(true)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_path_functions() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 parent = fso.GetParentFolderName(\"/a/b/c.txt\")\n\
                 fname = fso.GetFileName(\"/a/b/c.txt\")\n\
                 ext = fso.GetExtensionName(\"/a/b/c.txt\")\n\
                 base = fso.GetBaseName(\"/a/b/c.txt\")\n\
                 built = fso.BuildPath(\"/a/b\", \"c.txt\")",
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("parent"),
            Some(&VBValue::String("/a/b".to_string()))
        );
        assert_eq!(
            ctx.get_variable("fname"),
            Some(&VBValue::String("c.txt".to_string()))
        );
        assert_eq!(
            ctx.get_variable("ext"),
            Some(&VBValue::String("txt".to_string()))
        );
        assert_eq!(
            ctx.get_variable("base"),
            Some(&VBValue::String("c".to_string()))
        );
        assert_eq!(
            ctx.get_variable("built"),
            Some(&VBValue::String("/a/b/c.txt".to_string()))
        );
    }

    #[test]
    fn test_fso_copyfile() {
        let src = tmp_path("copy_src.txt");
        let dst = tmp_path("copy_dst.txt");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::write(&src, "Copy test content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CopyFile \"{}\", \"{}\", True",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&dst).exists());
        let content = std::fs::read_to_string(&dst).unwrap();
        assert_eq!(content, "Copy test content");

        cleanup_path(&src);
        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_movefile() {
        let src = tmp_path("move_src.txt");
        let dst = tmp_path("move_dst.txt");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::write(&src, "Move test content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.MoveFile \"{}\", \"{}\"",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&src).exists());
        assert!(std::path::Path::new(&dst).exists());
        let content = std::fs::read_to_string(&dst).unwrap();
        assert_eq!(content, "Move test content");

        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_deletefile() {
        let path = tmp_path("delete_me.txt");
        cleanup_path(&path);
        std::fs::write(&path, "Delete me").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.DeleteFile \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&path).exists());
    }

    #[test]
    fn test_fso_create_delete_folder() {
        let path = tmp_path("test_folder");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Create folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CreateFolder \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert!(std::path::Path::new(&path).is_dir());

        // Check FolderExists
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     exists = fso.FolderExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert_eq!(ctx.get_variable("exists"), Some(&VBValue::Boolean(true)));

        // Delete folder
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.DeleteFolder \"{}\"",
                    path
                ),
                &mut ctx,
            )
            .unwrap();
        assert!(!std::path::Path::new(&path).exists());

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_fileobject_properties() {
        let path = tmp_path("fileobj.txt");
        cleanup_path(&path);
        std::fs::write(&path, "FileObject test").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     name = f.Name\n\
                     size = f.Size",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        // FileObject.Name returns the file name from the path
        let path_obj = std::path::Path::new(&path);
        let expected_name = path_obj.file_name().unwrap().to_str().unwrap().to_string();

        assert_eq!(
            ctx.get_variable("name"),
            Some(&VBValue::String(expected_name))
        );
        assert_eq!(ctx.get_variable("size"), Some(&VBValue::Number(15.0)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folderobject_properties() {
        let path = tmp_path("folderobj");
        cleanup_path(&path);
        std::fs::create_dir_all(&path).unwrap();
        std::fs::write(format!("{}/test.txt", &path), "hello").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     fname = fld.Name\n\
                     isRoot = fld.IsRootFolder",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let path_obj = std::path::Path::new(&path);
        let expected_name = path_obj.file_name().unwrap().to_str().unwrap().to_string();
        assert_eq!(
            ctx.get_variable("fname"),
            Some(&VBValue::String(expected_name))
        );
        assert_eq!(ctx.get_variable("isRoot"), Some(&VBValue::Boolean(false)));

        // Test Files collection
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     files = fld.Files\n\
                     fileCount = 0\n\
                     For Each f In files\n\
                         fileCount = fileCount + 1\n\
                     Next",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("fileCount"), Some(&VBValue::Number(1.0)));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_getabsolutepathname() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 absPath = fso.GetAbsolutePathName(\"test.txt\")",
                &mut ctx,
            )
            .unwrap();

        let abs = ctx.get_variable("absPath");
        assert!(abs.is_some());
        if let Some(VBValue::String(s)) = abs {
            assert!(s.ends_with("test.txt"));
            assert!(std::path::Path::new(s).is_absolute());
        } else {
            panic!("Expected String");
        }
    }

    #[test]
    fn test_fso_copyfolder() {
        let src = tmp_path("copyfolder_src");
        let dst = tmp_path("copyfolder_dst");
        cleanup_path(&src);
        cleanup_path(&dst);

        std::fs::create_dir_all(&src).unwrap();
        std::fs::write(format!("{}/a.txt", &src), "file a").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     fso.CopyFolder \"{}\", \"{}\", True",
                    src, dst
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&dst).is_dir());
        assert!(std::path::Path::new(&format!("{}/a.txt", &dst)).exists());

        cleanup_path(&src);
        cleanup_path(&dst);
    }

    #[test]
    fn test_fso_writeblanklines() {
        let path = tmp_path("blanklines.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.CreateTextFile(\"{}\", True)\n\
                     ts.WriteLine \"First\"\n\
                     ts.WriteBlankLines 2\n\
                     ts.WriteLine \"Last\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let content = std::fs::read_to_string(&path).unwrap();
        let lines: Vec<&str> = content.lines().collect();
        assert_eq!(lines.len(), 4);
        assert_eq!(lines[0], "First");
        assert_eq!(lines[1], "");
        assert_eq!(lines[2], "");
        assert_eq!(lines[3], "Last");

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_file_delete_method() {
        let path = tmp_path("filedelete.txt");
        cleanup_path(&path);
        std::fs::write(&path, "delete via method").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     f.Delete",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(!std::path::Path::new(&path).exists());
    }

    #[test]
    fn test_fso_getspecialfolder() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // TemporaryFolder (2)
        interp
            .execute(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 tmpFolder = fso.GetSpecialFolder(2)",
                &mut ctx,
            )
            .unwrap();

        let tmp = ctx.get_variable("tmpFolder");
        assert!(tmp.is_some());
        if let Some(VBValue::String(s)) = tmp {
            let p = std::path::Path::new(s);
            assert!(p.is_absolute());
        } else {
            panic!("Expected string path");
        }
    }

    #[test]
    fn test_fso_createobject_invalid_progid() {
        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        let result = interp.execute(
            "Set obj = CreateObject(\"Some.Nonexistent.Object\")",
            &mut ctx,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_fso_file_exists_false() {
        let path = tmp_path("nonexistent_file.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     exists = fso.FileExists(\"{}\")",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(ctx.get_variable("exists"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_fso_getfile_notfound_error() {
        let path = tmp_path("nonexistent_file_for_get.txt");
        cleanup_path(&path);

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        let result = interp.execute(
            &format!(
                "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                 Set f = fso.GetFile(\"{}\")",
                path
            ),
            &mut ctx,
        );
        assert!(result.is_err());
    }

    #[test]
    fn test_fso_textstream_append() {
        let path = tmp_path("append.txt");
        cleanup_path(&path);
        std::fs::write(&path, "Initial\n").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        // Append mode (8)
        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set ts = fso.OpenTextFile(\"{}\", 8, True)\n\
                     ts.WriteLine \"Appended\"\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        let content = std::fs::read_to_string(&path).unwrap();
        assert!(content.contains("Appended"));

        cleanup_path(&path);
    }

    #[test]
    fn test_fso_folder_createtextfile() {
        let folder = tmp_path("folder_create_file");
        let file_path = format!("{}/newfile.txt", &folder);
        cleanup_path(&folder);
        std::fs::create_dir_all(&folder).unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set fld = fso.GetFolder(\"{}\")\n\
                     Set ts = fld.CreateTextFile(\"newfile.txt\", True)\n\
                     ts.WriteLine \"Created via Folder\"\n\
                     ts.Close",
                    folder
                ),
                &mut ctx,
            )
            .unwrap();

        assert!(std::path::Path::new(&file_path).exists());
        let content = std::fs::read_to_string(&file_path).unwrap();
        assert_eq!(content.trim(), "Created via Folder");

        cleanup_path(&folder);
    }

    #[test]
    fn test_fso_file_openastextstream() {
        let path = tmp_path("openas.txt");
        cleanup_path(&path);
        std::fs::write(&path, "OpenAsTextStream content").unwrap();

        let mut ctx = ExecutionContext::new();
        crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
        let interp = VBScriptInterpreter;

        interp
            .execute(
                &format!(
                    "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n\
                     Set f = fso.GetFile(\"{}\")\n\
                     Set ts = f.OpenAsTextStream(1)\n\
                     content = ts.ReadAll()\n\
                     ts.Close",
                    path
                ),
                &mut ctx,
            )
            .unwrap();

        assert_eq!(
            ctx.get_variable("content"),
            Some(&VBValue::String("OpenAsTextStream content".to_string()))
        );

        cleanup_path(&path);
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
        assert_eq!(ctx.scope.err_number, 42.0);
        assert_eq!(ctx.scope.err_description, "custom error");
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
        assert!(ctx.scope.err_number != 0.0);
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
        assert_eq!(ctx.scope.err_number, 5.0);
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
            Some(&VBValue::String("hello there there".to_string()))
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
            Some(&VBValue::String("hello there world".to_string()))
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
            assert_eq!(match0_value, VBValue::String("123".to_string()));
            assert_eq!(match1_value, VBValue::String("456".to_string()));
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
            Some(&VBValue::String("test".to_string()))
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
                "Provider=SQLOLEDB;Data Source=server".to_string()
            ))
        );
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
            Some(&VBValue::String("hi   ".to_string()))
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
            Some(&VBValue::String("hel".to_string()))
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
            Some(&VBValue::String("   hi".to_string()))
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
            Some(&VBValue::String("hel".to_string()))
        );
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

    // ===== HTTP INTEGRATION TESTS =====

    use std::sync::atomic::{AtomicU64, Ordering};
    use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};

    static TEST_COUNTER: AtomicU64 = AtomicU64::new(1);

    fn tmp_asp_dir() -> std::path::PathBuf {
        // Use the test function name if detected via backtrace. Fallback to counter.
        let counter = TEST_COUNTER.fetch_add(1, Ordering::Relaxed);
        let dir =
            std::env::temp_dir().join(format!("asperger_test_{}_{}", std::process::id(), counter));
        let _ = std::fs::create_dir_all(&dir);
        dir
    }

    fn write_asp(dir: &std::path::Path, name: &str, content: &str) {
        std::fs::write(
            dir.join(name),
            format!("<%@ LANGUAGE=VBScript %>{}", content),
        )
        .unwrap();
    }

    fn cleanup_dir(dir: &std::path::Path) {
        if dir.exists() {
            for entry in std::fs::read_dir(dir).unwrap() {
                if let Ok(e) = entry {
                    let _ = std::fs::remove_file(e.path());
                }
            }
            let _ = std::fs::remove_dir(dir);
        }
    }

    async fn serve_and_get(asp_dir: &str, request: &str) -> String {
        let config = crate::asp::config::Config {
            host: "127.0.0.1".to_string(),
            port: 0,
            folder: asp_dir.to_string(),
            program: None,
            enable_directory_listing: false,
        };
        let server = crate::asp::server::AspServer::new(config);
        let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();

        let folder = asp_dir.to_string();
        let handler = Arc::clone(&server.handler_chain);
        let store = Arc::clone(&server.store);
        let (shutdown_tx, mut shutdown_rx) = tokio::sync::oneshot::channel::<()>();

        let _handle = tokio::spawn(async move {
            loop {
                tokio::select! {
                    result = listener.accept() => {
                        let (stream, _) = result.unwrap();
                        let handler = Arc::clone(&handler);
                        let store = Arc::clone(&store);
                        let folder = folder.clone();
                        tokio::spawn(async move {
                            let mut stream = stream;
                            let default_doc = "index.asp".to_string();
                            let _ = crate::asp::server::AspServer::handle_connection(
                                &handler, &mut stream, &folder, &default_doc, &store, false,
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
}
