use std::iter::Peekable;
use std::str::Chars;
use std::sync::Arc;

use super::types::kw;
use super::{Token, TokenType};

pub struct Tokenizer<'a> {
    input: Peekable<Chars<'a>>,
    current_line: usize,
    current_column: usize,
    current_pos: usize,
}

impl<'a> Tokenizer<'a> {
    pub fn new(code: &'a str) -> Self {
        Tokenizer {
            input: code.chars().peekable(),
            current_line: 1,
            current_column: 1,
            current_pos: 0,
        }
    }

    fn tok(tt: TokenType, value: String) -> Token {
        Token { token_type: tt, value: Arc::from(value) }
    }

    pub fn tokenize(code: &'a str) -> Vec<Token> {
        let mut tokenizer = Tokenizer::new(code);
        let mut tokens = Vec::new();

        while let Some(token) = tokenizer.next_token() {
            if token.token_type == TokenType::EOF {
                tokens.push(token);
                break;
            }
            if token.token_type != TokenType::WhiteSpace {
                tokens.push(token);
            }
        }

        tokens
    }

    fn next_token(&mut self) -> Option<Token> {
        while let Some(&c) = self.input.peek() {
            match c {
                ' ' | '\t' => {
                    self.consume_whitespace();
                    continue;
                }
                '\n' | '\r' => {
                    return Some(self.handle_newline());
                }
                '"' => {
                    return Some(self.tokenize_string());
                }
                '#' => {
                    return Some(self.tokenize_date());
                }
                '0'..='9' => {
                    return Some(self.tokenize_number());
                }
                '\'' => {
                    return Some(self.tokenize_comment());
                }
                '_' => {
                    if self.is_line_continuation() {
                        self.consume_line_continuation();
                        continue;
                    }
                    return Some(self.tokenize_identifier());
                }
                c if self.is_identifier_start(c) => {
                    return Some(self.tokenize_identifier());
                }
                c if self.is_operator_char(c) => {
                    return Some(self.tokenize_operator());
                }
                _ => {
                    self.advance();
                    continue;
                }
            }
        }

        Some(Self::tok(TokenType::EOF, String::new()))
    }

    fn consume_whitespace(&mut self) {
        while let Some(&c) = self.input.peek() {
            if !c.is_whitespace() || c == '\n' || c == '\r' {
                break;
            }
            self.advance();
        }
    }

    fn handle_newline(&mut self) -> Token {
        let mut value = String::new();

        while let Some(&c) = self.input.peek() {
            if c != '\n' && c != '\r' {
                break;
            }
            value.push(c);
            self.advance();
            if c == '\n' {
                self.current_line += 1;
                self.current_column = 1;
            }
        }
        Self::tok(TokenType::NewLine, value)
    }

    fn tokenize_string(&mut self) -> Token {
        let mut value = String::new();

        self.advance();

        while let Some(&c) = self.input.peek() {
            self.advance();
            if c == '"' {
                if self.input.peek() == Some(&'"') {
                    value.push(c);
                    self.advance();
                } else {
                    break;
                }
            } else {
                value.push(c);
            }
        }

        Self::tok(TokenType::StringLiteral, value)
    }

    fn tokenize_number(&mut self) -> Token {
        let mut value = String::new();
        let mut is_float = false;
        let mut is_hex = false;
        let mut is_oct = false;

        if self.input.peek() == Some(&'&') {
            self.advance();
            value.push('&');
            if let Some(&next) = self.input.peek() {
                if next == 'H' || next == 'h' {
                    is_hex = true;
                    self.advance();
                    value.push('H');
                } else {
                    is_oct = true;
                }
            }
        }

        while let Some(&c) = self.input.peek() {
            match c {
                '0'..='9' | 'A'..='F' | 'a'..='f' if is_hex => {
                    value.push(c);
                    self.advance();
                }
                '0'..='7' if is_oct => {
                    value.push(c);
                    self.advance();
                }
                '0'..='9' => {
                    value.push(c);
                    self.advance();
                }
                '.' if !is_float && !is_hex && !is_oct => {
                    is_float = true;
                    value.push(c);
                    self.advance();
                }
                'E' | 'e' if !is_hex && !is_oct => {
                    is_float = true;
                    value.push(c);
                    self.advance();
                    if let Some(&next) = self.input.peek() {
                        if next == '+' || next == '-' {
                            value.push(next);
                            self.advance();
                        }
                    }
                }
                _ => break,
            }
        }

        let token_type = if is_float {
            TokenType::FloatLiteral
        } else if is_hex {
            TokenType::HexLiteral
        } else if is_oct {
            TokenType::OctLiteral
        } else {
            TokenType::IntegerLiteral
        };

        Self::tok(token_type, value)
    }

    fn tokenize_identifier(&mut self) -> Token {
        let mut value = String::new();

        while let Some(&c) = self.input.peek() {
            if self.is_identifier_char(c) {
                value.push(c);
                self.advance();
            } else {
                break;
            }
        }

        Self::tok(kw(&value), value)
    }

    fn is_identifier_start(&self, c: char) -> bool {
        c.is_alphabetic() || c == '_' || c == '['
    }

    fn is_identifier_char(&self, c: char) -> bool {
        c.is_alphanumeric() || c == '_' || c == ']'
    }

    fn is_operator_char(&self, c: char) -> bool {
        matches!(
            c,
            '+' | '-'
                | '*'
                | '/'
                | '\\'
                | '^'
                | '&'
                | '='
                | '.'
                | ','
                | ':'
                | '('
                | ')'
                | '>'
                | '<'
        )
    }

    fn advance(&mut self) {
        self.input.next();
        self.current_pos += 1;
        self.current_column += 1;
    }

    fn is_line_continuation(&mut self) -> bool {
        if let Some(&'_') = self.input.peek() {
            let mut chars = self.input.clone();
            chars.next();
            while let Some(&c) = chars.peek() {
                if c.is_whitespace() {
                    if c == '\n' || c == '\r' {
                        return true;
                    }
                    chars.next();
                } else {
                    break;
                }
            }
        }
        false
    }

    fn consume_line_continuation(&mut self) {
        self.advance();
        self.consume_whitespace();
        self.handle_newline();
    }

    fn tokenize_comment(&mut self) -> Token {
        let mut value = String::new();

        self.advance();

        while let Some(&c) = self.input.peek() {
            if c == '\n' || c == '\r' {
                break;
            }
            value.push(c);
            self.advance();
        }

        Self::tok(TokenType::Comment, value)
    }

    fn tokenize_date(&mut self) -> Token {
        let mut value = String::new();

        self.advance();

        while let Some(&c) = self.input.peek() {
            if c == '#' {
                self.advance();
                break;
            }
            value.push(c);
            self.advance();
        }

        Self::tok(TokenType::DateLiteral, value)
    }

    fn tokenize_operator(&mut self) -> Token {
        let mut value = String::new();

        if let Some(&c) = self.input.peek() {
            value.push(c);
            self.advance();

            match c {
                '+' => Self::tok(TokenType::Plus, value),
                '-' => Self::tok(TokenType::Minus, value),
                '*' => Self::tok(TokenType::Multiply, value),
                '/' => Self::tok(TokenType::Divide, value),
                '\\' => Self::tok(TokenType::IntDivide, value),
                '^' => Self::tok(TokenType::Power, value),
                '&' => {
                    if let Some(&next) = self.input.peek() {
                        if next == 'H' || next == 'h' {
                            value.push('H');
                            self.advance();
                            while let Some(&c) = self.input.peek() {
                                if c.is_ascii_hexdigit() {
                                    value.push(c);
                                    self.advance();
                                } else {
                                    break;
                                }
                            }
                            return Self::tok(TokenType::HexLiteral, value);
                        }
                        if next.is_ascii_digit() || next == 'O' || next == 'o' {
                            if next == 'O' || next == 'o' {
                                value.push(next);
                                self.advance();
                            }
                            while let Some(&c) = self.input.peek() {
                                if c.is_ascii_digit() && c != '8' && c != '9' {
                                    value.push(c);
                                    self.advance();
                                } else {
                                    break;
                                }
                            }
                            return Self::tok(TokenType::OctLiteral, value);
                        }
                    }
                    Self::tok(TokenType::Concat, value)
                }
                '=' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance();
                        return Self::tok(TokenType::Equal, value);
                    }
                    Self::tok(TokenType::Assign, value)
                }
                '.' => Self::tok(TokenType::Dot, value),
                ',' => Self::tok(TokenType::Comma, value),
                ':' => Self::tok(TokenType::Colon, value),
                '(' => Self::tok(TokenType::LeftParen, value),
                ')' => Self::tok(TokenType::RightParen, value),
                '>' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance();
                        return Self::tok(TokenType::GreaterEqual, value);
                    }
                    Self::tok(TokenType::GreaterThan, value)
                }
                '<' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance();
                        return Self::tok(TokenType::LessEqual, value);
                    } else if self.input.peek() == Some(&'>') {
                        value.push('>');
                        self.advance();
                        return Self::tok(TokenType::NotEqual, value);
                    }
                    Self::tok(TokenType::LessThan, value)
                }
                _ => {
                    self.advance();
                    Self::tok(TokenType::Invalid, value)
                }
            }
        } else {
            Self::tok(TokenType::EOF, value)
        }
    }
}
