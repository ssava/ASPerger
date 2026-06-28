//! Tokenizer / lexer for the VBScript language. Converts source text
//! into a sequence of `Token` values with associated `TokenType` tags.

use std::iter::Peekable;
use std::str::Chars;
use std::sync::Arc;

#[derive(Debug, PartialEq, Clone, Copy)]
pub enum TokenType {
    // Keywords
    Class,
    Function,
    Sub,
    Dim,
    If,
    Then,
    Else,
    ElseIf,
    End,
    For,
    Next,
    Do,
    Loop,
    While,
    WEnd,
    Select,
    Case,
    With,
    Set,
    New,
    True,
    False,
    Nothing,
    Null,
    Empty,
    And,
    Or,
    Not,
    Mod,
    Is,
    Eqv,
    Imp,
    To,
    Step,
    ReDim,
    Preserve,
    Property,
    Get,
    Let,
    Public,
    Private,
    Const,

    // Operators
    Plus,
    Minus,
    Multiply,
    Divide,
    IntDivide,
    Power,
    Concat,
    Assign,
    Dot,
    Comma,
    Colon,
    LeftParen,
    RightParen,
    GreaterThan,
    LessThan,
    GreaterEqual,
    LessEqual,
    NotEqual,
    Equal,

    // Literals
    Identifier,
    StringLiteral,
    IntegerLiteral,
    FloatLiteral,
    DateLiteral,
    HexLiteral,
    OctLiteral,

    // Other
    NewLine,
    Comment,
    WhiteSpace,
    Invalid,
    EOF,
}

/// A single lexical token produced by the `Tokenizer`.
#[derive(Debug, Clone)]
pub struct Token {
    /// The kind of token (keyword, operator, literal, etc.).
    pub token_type: TokenType,
    /// The source text of this token (shared via `Arc<str>`).
    pub value: Arc<str>,
}

/// VBScript lexer.  Converts source text into a sequence of `Token` values.
///
/// Handles string literals (with `""` escaping), numeric literals (integer,
/// float, hex, octal), date literals (`#...#`), identifiers, keywords
/// (case-insensitive), operators, comments (`'` / `REM`), and line
/// continuations (`_` + newline).
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

        self.advance(); // consume opening quote

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

        // Check for hex/oct prefix
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

        Self::tok(self.get_keyword_type(&value), value)
    }

    fn kw(word: &str) -> TokenType {
        match word.to_uppercase().as_str() {
            "CLASS" => TokenType::Class,
            "FUNCTION" => TokenType::Function,
            "SUB" => TokenType::Sub,
            "DIM" => TokenType::Dim,
            "IF" => TokenType::If,
            "THEN" => TokenType::Then,
            "ELSE" => TokenType::Else,
            "ELSEIF" => TokenType::ElseIf,
            "END" => TokenType::End,
            "FOR" => TokenType::For,
            "NEXT" => TokenType::Next,
            "DO" => TokenType::Do,
            "LOOP" => TokenType::Loop,
            "WHILE" => TokenType::While,
            "WEND" => TokenType::WEnd,
            "SELECT" => TokenType::Select,
            "CASE" => TokenType::Case,
            "WITH" => TokenType::With,
            "SET" => TokenType::Set,
            "NEW" => TokenType::New,
            "TRUE" => TokenType::True,
            "FALSE" => TokenType::False,
            "NOTHING" => TokenType::Nothing,
            "NULL" => TokenType::Null,
            "EMPTY" => TokenType::Empty,
            "AND" => TokenType::And,
            "OR" => TokenType::Or,
            "NOT" => TokenType::Not,
            "MOD" => TokenType::Mod,
            "IS" => TokenType::Is,
            "EQV" => TokenType::Eqv,
            "IMP" => TokenType::Imp,
            "TO" => TokenType::To,
            "STEP" => TokenType::Step,
            "REDIM" => TokenType::ReDim,
            "PRESERVE" => TokenType::Preserve,
            "PROPERTY" => TokenType::Property,
            "GET" => TokenType::Get,
            "LET" => TokenType::Let,
            "PUBLIC" => TokenType::Public,
            "PRIVATE" => TokenType::Private,
            "CONST" => TokenType::Const,
            _ => TokenType::Identifier,
        }
    }

    fn get_keyword_type(&self, word: &str) -> TokenType {
        Self::kw(word)
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
            chars.next(); // consume '_'
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
        self.advance(); // consume '_'
        self.consume_whitespace();
        self.handle_newline();
    }

    fn tokenize_comment(&mut self) -> Token {
        let mut value = String::new();

        self.advance(); // consume the comment character

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

        self.advance(); // consume opening #

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
