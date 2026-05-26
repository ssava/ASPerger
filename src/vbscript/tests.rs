#[cfg(test)]
mod tests {
    use crate::asp::parser::AspParser;
    use crate::vbscript::expr::{evaluate, parse_expression, BinOp, Expr, UnaryOp};
    use crate::vbscript::syntax::{Assignment, Dim, ResponseWrite, VBSyntax};
    use crate::vbscript::{ExecutionContext, TokenType, Tokenizer, VBValue};

    // ===== TOKENIZER =====

    #[test]
    fn test_tokenizer_simple_assignment() {
        let tokens = Tokenizer::tokenize("x = 5");
        assert_eq!(tokens.len(), 4);
        assert_eq!(tokens[0].token_type, TokenType::Identifier);
        assert_eq!(tokens[0].value, "x");
        assert_eq!(tokens[1].token_type, TokenType::Assign);
        assert_eq!(tokens[2].token_type, TokenType::IntegerLiteral);
        assert_eq!(tokens[2].value, "5");
        assert_eq!(tokens[3].token_type, TokenType::EOF);
    }

    #[test]
    fn test_tokenizer_keywords() {
        let tokens = Tokenizer::tokenize("If Then Else ElseIf End For Next Do Loop While Wend");
        let expected = [
            TokenType::If, TokenType::Then, TokenType::Else, TokenType::ElseIf,
            TokenType::End, TokenType::For, TokenType::Next, TokenType::Do,
            TokenType::Loop, TokenType::While, TokenType::WEnd,
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
            TokenType::And, TokenType::Or, TokenType::Not, TokenType::Mod,
            TokenType::Is, TokenType::Eqv, TokenType::Imp, TokenType::To, TokenType::Step,
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
        assert_eq!(tokens[0].value, r#"he"llo"#);
    }

    #[test]
    fn test_tokenizer_comment() {
        let tokens = Tokenizer::tokenize("' this is a comment");
        assert_eq!(tokens[0].token_type, TokenType::Comment);
        assert_eq!(tokens[0].value, " this is a comment");
        assert_eq!(tokens[1].token_type, TokenType::EOF);
    }

    #[test]
    fn test_tokenizer_hex_literal() {
        let tokens = Tokenizer::tokenize("&HFF");
        assert_eq!(tokens[0].token_type, TokenType::HexLiteral);
        assert_eq!(tokens[0].value, "&HFF");
    }

    #[test]
    fn test_tokenizer_oct_literal() {
        let tokens = Tokenizer::tokenize("&77");
        assert_eq!(tokens[0].token_type, TokenType::OctLiteral);
        assert_eq!(tokens[0].value, "&77");
    }

    #[test]
    fn test_tokenizer_float() {
        let tokens = Tokenizer::tokenize("3.14");
        assert_eq!(tokens[0].token_type, TokenType::FloatLiteral);
        assert_eq!(tokens[0].value, "3.14");
    }

    #[test]
    fn test_tokenizer_scientific_notation() {
        let tokens = Tokenizer::tokenize("1.5e2");
        assert_eq!(tokens[0].token_type, TokenType::FloatLiteral);
        assert_eq!(tokens[0].value, "1.5e2");
    }

    #[test]
    fn test_tokenizer_date_literal() {
        let tokens = Tokenizer::tokenize("#2024-01-15#");
        assert_eq!(tokens[0].token_type, TokenType::DateLiteral);
        assert_eq!(tokens[0].value, "2024-01-15");
    }

    #[test]
    fn test_tokenizer_comparison_equal_vs_assign() {
        let tokens = Tokenizer::tokenize("x = 5");
        assert_eq!(tokens[1].token_type, TokenType::Assign);
        let tokens2 = Tokenizer::tokenize("x == 5");
        assert_eq!(tokens2[1].token_type, TokenType::Equal);
        assert_eq!(tokens2[1].value, "==");
    }

    #[test]
    fn test_tokenizer_operators() {
        let tokens = Tokenizer::tokenize("+ - * / \\ ^ & ( ) , . : > < >= <= <>");
        let expected = [
            TokenType::Plus, TokenType::Minus, TokenType::Multiply, TokenType::Divide,
            TokenType::IntDivide, TokenType::Power, TokenType::Concat,
            TokenType::LeftParen, TokenType::RightParen, TokenType::Comma, TokenType::Dot,
            TokenType::Colon, TokenType::GreaterThan, TokenType::LessThan,
            TokenType::GreaterEqual, TokenType::LessEqual, TokenType::NotEqual,
        ];
        for (i, expected_tt) in expected.iter().enumerate() {
            assert_eq!(tokens[i].token_type, *expected_tt, "mismatch at token {}: got {:?} value={:?}", i, tokens[i].token_type, tokens[i].value);
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
    fn test_evaluate_undefined_variable_error() {
        let context = ExecutionContext::new();
        let expr = Expr::Variable("undefined".into());
        assert!(evaluate(&expr, &context).is_err());
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
    fn test_evaluate_power() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Pow,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(8.0));
    }

    #[test]
    fn test_evaluate_int_divide() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::IntDiv,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(3.0));
    }

    #[test]
    fn test_evaluate_modulo() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(10.0))),
            op: BinOp::Mod,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(1.0));
    }

    #[test]
    fn test_evaluate_unary_neg() {
        let context = ExecutionContext::new();
        let expr = Expr::UnaryOp {
            op: UnaryOp::Neg,
            expr: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(-5.0));
    }

    #[test]
    fn test_evaluate_neg_of_negative() {
        let context = ExecutionContext::new();
        let expr = Expr::UnaryOp {
            op: UnaryOp::Neg,
            expr: Box::new(Expr::Literal(VBValue::Number(-3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(3.0));
    }

    // ===== EXPRESSION EVALUATOR — STRING CONCAT =====

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
    fn test_evaluate_concat_number_coercion() {
        let context = ExecutionContext::new();
        // "a" & 42 → "a42"
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::Number(42.0))),
        };
        assert_eq!(
            evaluate(&expr, &context).unwrap(),
            VBValue::String("a42".into())
        );
    }

    // ===== EXPRESSION EVALUATOR — BOOLEAN / LOGICAL =====

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
            op: UnaryOp::Not,
            expr: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_xor() {
        let context = ExecutionContext::new();
        // true XOR true = false
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Xor,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(false));

        // true XOR false = true
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Xor,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(evaluate(&expr2, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_eqv() {
        let context = ExecutionContext::new();
        // true Eqv true = true
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Eqv,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));
        // true Eqv false = false
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Eqv,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(evaluate(&expr2, &context).unwrap(), VBValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_imp() {
        let context = ExecutionContext::new();
        // true Imp false = false
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Imp,
            right: Box::new(Expr::Literal(VBValue::Boolean(false))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(false));
        // false Imp anything = true
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(false))),
            op: BinOp::Imp,
            right: Box::new(Expr::Literal(VBValue::Boolean(true))),
        };
        assert_eq!(evaluate(&expr2, &context).unwrap(), VBValue::Boolean(true));
    }

    // ===== EXPRESSION EVALUATOR — COMPARISON =====

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
    fn test_evaluate_comparison_ne() {
        let context = ExecutionContext::new();
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Ne,
            right: Box::new(Expr::Literal(VBValue::Number(3.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));

        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(5.0))),
            op: BinOp::Ne,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr2, &context).unwrap(), VBValue::Boolean(false));
    }

    #[test]
    fn test_evaluate_comparison_lt_gt() {
        let context = ExecutionContext::new();
        let lt = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(1.0))),
            op: BinOp::Lt,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&lt, &context).unwrap(), VBValue::Boolean(true));

        let gt = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Gt,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(evaluate(&gt, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_comparison_le_ge() {
        let context = ExecutionContext::new();
        let le = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(2.0))),
            op: BinOp::Le,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&le, &context).unwrap(), VBValue::Boolean(true));

        let ge = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Number(3.0))),
            op: BinOp::Ge,
            right: Box::new(Expr::Literal(VBValue::Number(2.0))),
        };
        assert_eq!(evaluate(&ge, &context).unwrap(), VBValue::Boolean(true));
    }

    #[test]
    fn test_evaluate_is_operator() {
        let context = ExecutionContext::new();
        // Is compares values — two nulls are equivalent
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Null)),
            op: BinOp::Is,
            right: Box::new(Expr::Literal(VBValue::Null)),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Boolean(true));
    }

    // ===== EXPRESSION EVALUATOR — TYPE COERCION =====

    #[test]
    fn test_evaluate_add_string_coercion() {
        let context = ExecutionContext::new();
        // Number + String → string concat in VBScript-semantics
        // Our implementation: if both are numeric, add; else concat
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        // String + Number → concat as strings
        let result = evaluate(&expr, &context).unwrap();
        assert_eq!(result, VBValue::String("a1".into()));
    }

    #[test]
    fn test_evaluate_empty_as_number() {
        let context = ExecutionContext::new();
        // Empty acts as 0 in numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Empty)),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_empty_as_string() {
        let context = ExecutionContext::new();
        // Empty acts as "" in string context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("x".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::Empty)),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::String("x".into()));
    }

    #[test]
    fn test_evaluate_empty() {
        let context = ExecutionContext::new();
        let expr = Expr::Literal(VBValue::Empty);
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Empty);
    }

    #[test]
    fn test_evaluate_null_as_number() {
        let context = ExecutionContext::new();
        // Null acts as 0 in numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Null)),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(5.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(5.0));
    }

    #[test]
    fn test_evaluate_boolean_as_number() {
        let context = ExecutionContext::new();
        // True = -1 in VBScript numeric context
        let expr = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(true))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(evaluate(&expr, &context).unwrap(), VBValue::Number(0.0));

        // False = 0
        let expr2 = Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::Boolean(false))),
            op: BinOp::Add,
            right: Box::new(Expr::Literal(VBValue::Number(1.0))),
        };
        assert_eq!(evaluate(&expr2, &context).unwrap(), VBValue::Number(1.0));
    }

    // ===== ASSIGNMENT =====

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
        let expr = Expr::Variable("a".into());
        let assignment = Assignment::new("x".into(), expr);
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(10.0)));
    }

    #[test]
    fn test_assignment_string() {
        let mut context = ExecutionContext::new();
        let assignment = Assignment::new("s".into(), Expr::Literal(VBValue::String("hello".into())));
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("s"), Some(VBValue::String("hello".into())));
    }

    #[test]
    fn test_assignment_boolean() {
        let mut context = ExecutionContext::new();
        let assignment = Assignment::new("b".into(), Expr::Literal(VBValue::Boolean(true)));
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("b"), Some(VBValue::Boolean(true)));
    }

    #[test]
    fn test_assignment_null() {
        let mut context = ExecutionContext::new();
        let assignment = Assignment::new("n".into(), Expr::Literal(VBValue::Null));
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("n"), Some(VBValue::Null));
    }

    #[test]
    fn test_assignment_concat_expression() {
        let mut context = ExecutionContext::new();
        context.set_variable("pre", VBValue::String("Hello ".into()));
        let assignment = Assignment::new("msg".into(), Expr::BinaryOp {
            left: Box::new(Expr::Variable("pre".into())),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::String("World".into()))),
        });
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("msg"), Some(VBValue::String("Hello World".into())));
    }

    // ===== RESPONSE.WRITE =====

    #[test]
    fn test_response_write_literal() {
        let mut context = ExecutionContext::new();
        let rw = ResponseWrite::new(Expr::Literal(VBValue::String("hello".into())));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "hello");
    }

    #[test]
    fn test_response_write_variable() {
        let mut context = ExecutionContext::new();
        context.set_variable("name", VBValue::String("World".into()));
        let rw = ResponseWrite::new(Expr::Variable("name".into()));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "World");
    }

    #[test]
    fn test_response_write_number() {
        let mut context = ExecutionContext::new();
        let rw = ResponseWrite::new(Expr::Literal(VBValue::Number(42.0)));
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "42");
    }

    #[test]
    fn test_response_write_expression() {
        let mut context = ExecutionContext::new();
        let rw = ResponseWrite::new(Expr::BinaryOp {
            left: Box::new(Expr::Literal(VBValue::String("a".into()))),
            op: BinOp::Concat,
            right: Box::new(Expr::Literal(VBValue::String("b".into()))),
        });
        rw.execute(&mut context).unwrap();
        assert_eq!(context.response_buffer, "ab");
    }

    // ===== DIM =====

    #[test]
    fn test_dim_initializes_to_empty() {
        let mut context = ExecutionContext::new();
        let dim = Dim::new(vec!["x".into()]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Empty));
    }

    #[test]
    fn test_dim_multiple_variables() {
        let mut context = ExecutionContext::new();
        let dim = Dim::new(vec!["a".into(), "b".into(), "c".into()]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("a"), Some(VBValue::Empty));
        assert_eq!(context.get_variable("b"), Some(VBValue::Empty));
        assert_eq!(context.get_variable("c"), Some(VBValue::Empty));
    }

    // ===== BLOCK STATEMENTS — IF =====

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
    fn test_if_elseif_chain() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // First ElseIf matches
        context.set_variable("x", VBValue::Number(2.0));
        interpreter.execute("If x = 1 Then\n    y = 10\nElseIf x = 2 Then\n    y = 20\nElseIf x = 3 Then\n    y = 30\nElse\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(20.0)));
    }

    #[test]
    fn test_if_elseif_falls_to_else() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(9.0));
        interpreter.execute("If x = 1 Then\n    y = 10\nElseIf x = 2 Then\n    y = 20\nElse\n    y = 99\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(99.0)));
    }

    #[test]
    fn test_if_multiple_blocks_in_sequence() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute(
            "If 1 = 1 Then\n    a = 10\nEnd If\nIf 2 = 2 Then\n    b = 20\nEnd If",
            &mut context,
        )
        .unwrap();
        assert_eq!(context.get_variable("a"), Some(VBValue::Number(10.0)));
        assert_eq!(context.get_variable("b"), Some(VBValue::Number(20.0)));
    }

    #[test]
    fn test_if_condition_with_and() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(5.0));
        interpreter.execute("If x > 1 And x < 10 Then y = 1\nEnd If", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(1.0)));
    }

    // ===== BLOCK STATEMENTS — FOR =====

    #[test]
    fn test_for_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim total\ntotal = 0\nFor i = 1 To 5\n    total = total + i\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(15.0)));
    }

    #[test]
    fn test_for_with_step() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim total\ntotal = 0\nFor i = 1 To 10 Step 2\n    total = total + i\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(25.0)));
    }

    #[test]
    fn test_for_negative_step() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim total\ntotal = 0\nFor i = 5 To 1 Step -1\n    total = total + i\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(15.0)));
    }

    #[test]
    fn test_for_zero_iterations() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // start > end with positive step → zero iterations
        interpreter.execute("Dim total\ntotal = 99\nFor i = 5 To 1\n    total = 0\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(99.0)));
    }

    #[test]
    fn test_for_empty_body() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("For i = 1 To 3\nNext", &mut context).unwrap();
        // Just shouldn't crash; counter should be 4 (past end)
        assert_eq!(context.get_variable("i"), Some(VBValue::Number(4.0)));
    }

    #[test]
    fn test_for_modifies_counter_in_body() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        // Modifying the loop counter inside the body
        interpreter.execute("Dim total\ntotal = 0\nFor i = 1 To 10\n    total = total + i\n    i = i + 1\nNext", &mut context).unwrap();
        // i increments by 1 in loop header, plus +1 in body = effective step of 2
        // So i goes: 1 → (header: 2, body: +1) → 3 → 4 → 5 → 6 → 7 → 8 → 9 → header:10 → body:11 → header:12 stops
        // The loop counter is managed by the For statement,
        // so modifying i in the body doesn't affect iteration count.
        assert_eq!(context.get_variable("total"), Some(VBValue::Number(55.0)));
    }

    // ===== BLOCK STATEMENTS — WHILE =====

    #[test]
    fn test_while_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("While x <= 3\n    x = x + 1\nWend", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(4.0)));
    }

    #[test]
    fn test_while_never_enters() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter.execute("While x < 5\n    x = 99\nWend", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(10.0)));
    }

    #[test]
    fn test_while_empty_body_does_not_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(0.0));
        interpreter.execute("While x > 0\n    x = 99\nWend", &mut context).unwrap();
        // Should not modify x since condition is false
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(0.0)));
    }

    // ===== BLOCK STATEMENTS — DO =====

    #[test]
    fn test_do_while() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("Do While x < 3\n    x = x + 1\nLoop", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(3.0)));
    }

    #[test]
    fn test_do_while_never_enters() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter.execute("Do While x < 5\n    x = 99\nLoop", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(10.0)));
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
    fn test_do_loop_while_post_test() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(1.0));
        interpreter.execute("Do\n    x = x + 1\nLoop While x < 3", &mut context).unwrap();
        // Post-test: executes body, then checks. So: x=2 (check 2<3), x=3 (check 3<3=false)
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(3.0)));
    }

    #[test]
    fn test_do_loop_until_post_test_runs_once() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        context.set_variable("x", VBValue::Number(10.0));
        interpreter.execute("Do\n    x = x + 1\nLoop Until x > 5", &mut context).unwrap();
        // Post-test: always runs at least once. x=11 (check 11>5=true)
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(11.0)));
    }

    // ===== NESTED BLOCKS =====

    #[test]
    fn test_nested_blocks() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Dim result\nresult = 0\nFor i = 1 To 3\n    If i > 1 Then\n        result = result + i\n    End If\nNext", &mut context).unwrap();
        assert_eq!(context.get_variable("result"), Some(VBValue::Number(5.0)));
    }

    #[test]
    fn test_deeply_nested_blocks() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute(
            "Dim out\nout = 0\nFor a = 1 To 2\n    For b = 1 To 2\n        If b = a Then\n            out = out + 1\n        End If\n    Next\nNext",
            &mut context,
        )
        .unwrap();
        // a=1,b=1: matches → out=1; a=1,b=2: no; a=2,b=1: no; a=2,b=2: matches → out=2
        assert_eq!(context.get_variable("out"), Some(VBValue::Number(2.0)));
    }

    // ===== COMMENTS =====

    #[test]
    fn test_comment_apostrophe_line() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("' this is a comment\nx = 1", &mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(1.0)));
    }

    #[test]
    fn test_comment_rem_keyword() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("Rem this is a comment\ny = 2", &mut context).unwrap();
        assert_eq!(context.get_variable("y"), Some(VBValue::Number(2.0)));
    }

    #[test]
    fn test_code_with_only_comments() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("' comment 1\n' comment 2\nRem comment 3", &mut context).unwrap();
        // Just shouldn't crash
    }

    // ===== ERROR HANDLING =====

    #[test]
    fn test_undefined_variable_in_expression() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("x = undefinedVar + 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_division_by_zero_in_if_condition() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("If 1 / 0 = 1 Then x = 1\nEnd If", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_if_without_then() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("If x = 1\nEnd If", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_for_without_next() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("For i = 1 To 5\nx = 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_while_without_wend() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("While x < 5\nx = x + 1", &mut context);
        assert!(result.is_err());
    }

    #[test]
    fn test_syntax_error_do_without_loop() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        let result = interpreter.execute("Do While x < 5\nx = x + 1", &mut context);
        assert!(result.is_err());
    }

    // ===== EMPTY / WHITESPACE CODE =====

    #[test]
    fn test_empty_code() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("", &mut context).unwrap();
        // No crash
    }

    #[test]
    fn test_whitespace_code() {
        let mut context = ExecutionContext::new();
        let interpreter = crate::vbscript::VBScriptInterpreter;
        interpreter.execute("   \n  \t  \n  ", &mut context).unwrap();
        // No crash
    }

    // ===== PRESERVED BEHAVIOR =====

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
            crate::asp::parser::AspBlock::Code(c) => assert_eq!(c, "Dim x"),
            _ => panic!("Expected Code block"),
        }
    }

    #[test]
    fn test_asp_parser_only_html() {
        let parser = AspParser::new("<html><body>Hello</body></html>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Html(h) => assert_eq!(h, "<html><body>Hello</body></html>"),
            _ => panic!("Expected Html block"),
        }
    }

    #[test]
    fn test_asp_parser_only_code() {
        let parser = AspParser::new("<%x = 1%>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Code(c) => assert_eq!(c, "x = 1"),
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
}
