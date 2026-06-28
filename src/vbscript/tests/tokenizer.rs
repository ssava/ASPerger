use super::*;

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
