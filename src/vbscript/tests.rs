#[cfg(test)]
mod tests {
    use crate::asp::parser::AspParser;
    use crate::vbscript::syntax::{Assignment, Dim, VBSyntax};
    use crate::vbscript::{ExecutionContext, TokenType, Tokenizer, VBValue};

    #[test]
    fn test_tokenizer_simple_assignment() {
        let tokens = Tokenizer::tokenize("x = 5");
        assert_eq!(tokens.len(), 4);
        assert_eq!(tokens[0].token_type, TokenType::Identifier);
        assert_eq!(tokens[0].value, "x");
        assert_eq!(tokens[1].token_type, TokenType::Assign);
        assert_eq!(tokens[1].value, "=");
        assert_eq!(tokens[2].token_type, TokenType::IntegerLiteral);
        assert_eq!(tokens[2].value, "5");
        assert_eq!(tokens[3].token_type, TokenType::EOF);
    }

    #[test]
    fn test_assignment_execution() {
        let mut context = ExecutionContext::new();
        let assignment = Assignment::new("x".into(), "42".into());
        assignment.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Number(42.0)));
    }

    #[test]
    fn test_dim_execution() {
        let mut context = ExecutionContext::new();
        let dim = Dim::new(vec!["x".into()]);
        dim.execute(&mut context).unwrap();
        assert_eq!(context.get_variable("x"), Some(VBValue::Null));
    }

    #[test]
    fn test_asp_parser_splits_code_blocks() {
        let parser = AspParser::new("<%code%>".to_string());
        let blocks = parser.parse();
        assert_eq!(blocks.len(), 1);
        match &blocks[0] {
            crate::asp::parser::AspBlock::Code(code) => assert_eq!(code, "code"),
            _ => panic!("Expected Code block"),
        }
    }
}
