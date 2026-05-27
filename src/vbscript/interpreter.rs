use std::vec::Vec;

use crate::vbscript::block;
use crate::vbscript::vbs_error::VBSError;
use crate::vbscript::vbobject::ErrObject;
use crate::vbscript::{Token, TokenType, Tokenizer, VBValue};
use crate::vbscript::ExecutionContext;

pub struct VBScriptInterpreter;

impl VBScriptInterpreter {
    pub fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let code = code.trim().to_string();

        let tokens = Tokenizer::tokenize(&code);
        if tokens.iter().all(|t| t.token_type == TokenType::EOF) {
            return Ok(());
        }

        let lines = self.group_tokens_into_lines(&tokens)?;

        if context.get_variable("ERR").is_none() {
            context.set_variable("ERR", VBValue::Object(Box::new(ErrObject::new())));
        }

        let blocks = block::parse_blocks(&lines)?;
        block::execute_blocks(&blocks, context)
    }

    fn group_tokens_into_lines(&self, tokens: &[Token]) -> Result<Vec<Vec<Token>>, VBSError> {
        let mut lines: Vec<Vec<Token>> = Vec::new();
        let mut current_line: Vec<Token> = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let token = &tokens[i];

            match &token.token_type {
                TokenType::NewLine => {
                    if !current_line.is_empty() {
                        if !Self::is_line_continuation(&current_line) {
                            lines.push(std::mem::take(&mut current_line));
                        }
                    }
                }

                TokenType::WhiteSpace => {
                    if Self::is_continuation_sequence(&tokens[i..]) {
                        i = Self::skip_continuation_sequence(&tokens[i..]);
                    } else {
                        if !current_line.is_empty() {
                            current_line.push(token.clone());
                        }
                    }
                }

                TokenType::Comment => {
                    current_line.push(token.clone());
                    while i + 1 < tokens.len()
                        && !matches!(tokens[i + 1].token_type, TokenType::NewLine)
                    {
                        i += 1;
                    }
                }

                TokenType::Colon => {
                    current_line.push(token.clone());

                    if !Self::is_in_string_literal(&current_line) {
                        lines.push(std::mem::take(&mut current_line));
                    }
                }

                TokenType::EOF => {
                    break;
                }

                _ => {
                    current_line.push(token.clone());
                }
            }

            i += 1;
        }

        if !current_line.is_empty() {
            lines.push(current_line);
        }

        let lines = lines
            .into_iter()
            .map(|line| Self::trim_whitespace_tokens(line))
            .filter(|line| !line.is_empty())
            .collect();

        Ok(lines)
    }

    fn is_line_continuation(line: &[Token]) -> bool {
        if let Some(last_token) = line.last() {
            matches!(last_token.token_type, TokenType::WhiteSpace) && last_token.value.contains('_')
        } else {
            false
        }
    }

    fn is_in_string_literal(tokens: &[Token]) -> bool {
        let mut in_string = false;
        for token in tokens {
            match token.token_type {
                TokenType::StringLiteral => in_string = !in_string,
                _ => {}
            }
        }
        in_string
    }

    fn is_continuation_sequence(tokens: &[Token]) -> bool {
        if tokens.is_empty() {
            return false;
        }

        match &tokens[0].token_type {
            TokenType::WhiteSpace => {
                tokens[0].value.contains('_')
                    && tokens
                        .iter()
                        .skip(1)
                        .take_while(|t| matches!(t.token_type, TokenType::WhiteSpace))
                        .any(|t| matches!(t.token_type, TokenType::NewLine))
            }
            _ => false,
        }
    }

    fn skip_continuation_sequence(tokens: &[Token]) -> usize {
        let mut i = 0;

        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        while i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::NewLine) {
            i += 1;
        }

        i
    }

    fn trim_whitespace_tokens(tokens: Vec<Token>) -> Vec<Token> {
        let mut result = tokens;

        let first_non_ws = result
            .iter()
            .position(|t| !matches!(t.token_type, TokenType::WhiteSpace))
            .unwrap_or(result.len());
        result.drain(..first_non_ws);

        while result
            .last()
            .map_or(false, |t| matches!(t.token_type, TokenType::WhiteSpace))
        {
            result.pop();
        }

        result
    }
}
