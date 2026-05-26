use std::vec::Vec;

use crate::vbscript::expr::{parse_expression, Expr};
use crate::vbscript::syntax::{ResponseWrite, VBSyntax};
use crate::vbscript::ExecutionContext;

use super::syntax::{Assignment, Dim};
use super::vbs_error::{VBSError, VBSErrorType};
use super::{Token, TokenType, Tokenizer};
use super::VBValue;

pub struct VBScriptInterpreter;

impl VBScriptInterpreter {
    /// Executes the provided VBScript code by interpreting each line.
    ///
    /// # Arguments
    /// * `code` - A string slice containing the VBScript code to execute.
    /// * `context` - A mutable reference to the execution context where variables and functions are stored.
    ///
    /// # Returns
    /// * `Ok(())` if the execution is successful.
    /// * `Err(String)` if there is a syntax or runtime error.
    pub fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let code = code.trim().to_string();

        // Tokenize the entire code first
        let tokens = Tokenizer::tokenize(&code);
        if tokens.iter().all(|t| t.token_type == TokenType::EOF) {
            return Ok(());
        }

        // Group tokens into logical lines, handling line continuations
        let lines = self.group_tokens_into_lines(&tokens)?;

        let blocks = crate::vbscript::block::parse_blocks(&lines)?;
        crate::vbscript::block::execute_blocks(&blocks, self, context)
    }

    /// Groups tokens into logical lines, handling line continuations and special VBScript syntax rules.
    ///
    /// # Arguments
    /// * `tokens` - A slice of tokens from the VBScript code
    ///
    /// # Returns
    /// * `Result<Vec<Vec<Token>>, VBSError>` - A vector of token vectors, each representing a logical line
    fn group_tokens_into_lines(&self, tokens: &[Token]) -> Result<Vec<Vec<Token>>, VBSError> {
        let mut lines: Vec<Vec<Token>> = Vec::new();
        let mut current_line: Vec<Token> = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let token = &tokens[i];

            match &token.token_type {
                TokenType::NewLine => {
                    // Check if we should add the current line
                    if !current_line.is_empty() {
                        // Check if the previous token was a line continuation
                        if !Self::is_line_continuation(&current_line) {
                            lines.push(current_line.clone());
                            current_line.clear();
                        }
                    }
                }

                TokenType::WhiteSpace => {
                    // Look ahead for line continuation
                    if Self::is_continuation_sequence(&tokens[i..]) {
                        // Skip the continuation sequence and following newline
                        i = Self::skip_continuation_sequence(&tokens[i..]);
                        // Don't clear current_line - continue accumulating tokens
                    } else {
                        // Normal whitespace - add it if it's not at the start of a line
                        if !current_line.is_empty() {
                            current_line.push(token.clone());
                        }
                    }
                }

                TokenType::Comment => {
                    // Add comment to current line
                    current_line.push(token.clone());
                    // Skip to end of line
                    while i + 1 < tokens.len()
                        && !matches!(tokens[i + 1].token_type, TokenType::NewLine)
                    {
                        i += 1;
                    }
                }

                TokenType::Colon => {
                    // Handle statement separator
                    current_line.push(token.clone());

                    // Check if this colon is inside a string literal
                    if !Self::is_in_string_literal(&current_line) {
                        lines.push(current_line.clone());
                        current_line = Vec::new();
                    }
                }

                _ => {
                    current_line.push(token.clone());
                }
            }

            i += 1;
        }

        // Add the last line if it's not empty
        if !current_line.is_empty() {
            lines.push(current_line);
        }

        // Post-process: trim whitespace tokens at start/end of each line
        let lines = lines
            .into_iter()
            .map(|line| Self::trim_whitespace_tokens(line))
            .filter(|line| !line.is_empty())
            .collect();

        Ok(lines)
    }

    /// Helper function to check if a sequence of tokens represents a line continuation
    fn is_line_continuation(line: &[Token]) -> bool {
        if let Some(last_token) = line.last() {
            matches!(last_token.token_type, TokenType::WhiteSpace) && last_token.value.contains('_')
        } else {
            false
        }
    }

    /// Helper function to check if we're inside a string literal
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

    /// Helper function to check if a sequence of tokens starts with a line continuation
    fn is_continuation_sequence(tokens: &[Token]) -> bool {
        if tokens.is_empty() {
            return false;
        }

        // Check for underscore followed by optional whitespace and newline
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

    /// Helper function to skip over a line continuation sequence
    /// Returns the new index after the sequence
    fn skip_continuation_sequence(tokens: &[Token]) -> usize {
        let mut i = 0;

        // Skip initial whitespace with underscore
        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        // Skip any additional whitespace
        while i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        // Skip the newline token
        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::NewLine) {
            i += 1;
        }

        i
    }

    /// Helper function to trim whitespace tokens from start and end of a line
    fn trim_whitespace_tokens(tokens: Vec<Token>) -> Vec<Token> {
        let mut result = tokens;

        // Trim from start
        while result
            .first()
            .map_or(false, |t| matches!(t.token_type, TokenType::WhiteSpace))
        {
            result.remove(0);
        }

        // Trim from end
        while result
            .last()
            .map_or(false, |t| matches!(t.token_type, TokenType::WhiteSpace))
        {
            result.pop();
        }

        result
    }

    /// Creates a syntax object from a sequence of tokens
    pub(crate) fn create_syntax_from_tokens(
        &self,
        tokens: &[Token],
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        if tokens.is_empty() {
            return Ok(None);
        }

        // Get the first non-whitespace token to determine the statement type
        let first_token = tokens
            .iter()
            .find(|t| t.token_type != TokenType::WhiteSpace)
            .ok_or_else(|| VBSErrorType::SyntaxError.into_error("Empty statement".to_string()))?;

        match first_token.token_type {
            TokenType::Dim => self.parse_dim_statement(tokens),
            TokenType::Set => self.parse_assignment_statement(tokens),

            _ => {
                // Try to parse as expression or assignment if no keyword is recognized
                self.parse_expression_or_assignment(tokens)
            }
        }
    }

    /// Helper function to convert a sequence of tokens back to their string representation
    fn tokens_to_string(tokens: &[Token]) -> String {
        tokens
            .iter()
            .map(|t| t.value.clone())
            .collect::<Vec<String>>()
            .join(" ")
    }

    fn parse_expression_or_assignment(
        &self,
        tokens: &[Token],
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        let non_ws: Vec<&Token> = tokens.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

        // Case 1: Response.Write expr
        if non_ws.len() >= 3
            && non_ws[0].value.eq_ignore_ascii_case("response")
            && non_ws[1].token_type == TokenType::Dot
            && non_ws[2].value.eq_ignore_ascii_case("write")
        {
            // Find where the expression starts in the original token list
            let mut expr_start = tokens.len();
            let mut found_write = false;
            for (i, tok) in tokens.iter().enumerate() {
                if tok.token_type != TokenType::WhiteSpace && tok.value.eq_ignore_ascii_case("write") {
                    found_write = true;
                    continue;
                }
                if found_write {
                    expr_start = i;
                    break;
                }
            }

            let expr = if expr_start < tokens.len() {
                parse_expression(&tokens[expr_start..])?
            } else {
                Expr::Literal(VBValue::Empty)
            };
            return Ok(Some(Box::new(ResponseWrite::new(expr))));
        }

        // Case 2: var = expr (bare assignment, no Set keyword)
        if non_ws.len() >= 2
            && non_ws[0].token_type == TokenType::Identifier
            && non_ws[1].token_type == TokenType::Assign
        {
            let var_name = non_ws[0].value.clone();
            let assign_idx = tokens.iter().position(|t| t.token_type == TokenType::Assign).unwrap();
            let expr = parse_expression(&tokens[assign_idx + 1..])?;
            return Ok(Some(Box::new(Assignment::new(var_name, expr))));
        }

        // Replicate old behavior for unrecognized commands
        let line_text = Self::tokens_to_string(tokens);
        Err(VBSErrorType::NotImplementedError
            .into_error(format!("Unrecognized command: {}", line_text)))
    }

    fn parse_dim_statement(&self, tokens: &[Token]) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        // Ensure the first token is the `Dim` keyword
        if tokens.is_empty() || tokens[0].token_type != TokenType::Dim {
            return Err(VBSErrorType::SyntaxError.into_error("Expected 'Dim' keyword".to_string()));
        }
    
        // Collect variable names
        let mut var_names = Vec::new();
        let mut i = 1; // Start after the `Dim` keyword
    
        while i < tokens.len() {
            // Skip whitespace
            if tokens[i].token_type == TokenType::WhiteSpace {
                i += 1;
                continue;
            }
    
            // Expect an identifier (variable name)
            if tokens[i].token_type != TokenType::Identifier {
                return Err(VBSErrorType::SyntaxError.into_error(
                    format!("Expected variable name, found: {}", tokens[i].value)
                ));
            }
    
            // Add the variable name to the list
            var_names.push(tokens[i].value.clone());
    
            // Move to the next token
            i += 1;
    
            // Check for a comma (indicating another variable)
            if i < tokens.len() && tokens[i].token_type == TokenType::Comma {
                i += 1; // Skip the comma
            } else {
                break; // No more variables
            }
        }
    
        // Ensure we have at least one variable name
        if var_names.is_empty() {
            return Err(VBSErrorType::SyntaxError.into_error("No variable names found in 'Dim' statement".to_string()));
        }
    
        // Return a `Dim` syntax object
        Ok(Some(Box::new(Dim::new(var_names))))
    }

    fn parse_assignment_statement(
        &self,
        tokens: &[Token],
    ) -> Result<Option<Box<dyn VBSyntax>>, VBSError> {
        if tokens.is_empty() {
            return Err(VBSErrorType::SyntaxError.into_error("Empty assignment statement".to_string()));
        }

        let is_set_assignment = tokens[0].token_type == TokenType::Set;
        let mut i = if is_set_assignment { 1 } else { 0 };

        // Skip leading whitespace
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }

        if i >= tokens.len() || tokens[i].token_type != TokenType::Identifier {
            return Err(VBSErrorType::SyntaxError.into_error(
                format!("Expected variable name, found: {:?}", tokens.get(i)),
            ));
        }

        let var_name = tokens[i].value.clone();
        i += 1;

        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }

        if i >= tokens.len() || tokens[i].token_type != TokenType::Assign {
            return Err(VBSErrorType::SyntaxError.into_error(
                format!("Expected '=', found: {:?}", tokens.get(i)),
            ));
        }
        i += 1;

        let expr = parse_expression(&tokens[i..])?;
        Ok(Some(Box::new(Assignment::new(var_name, expr))))
    }


}
