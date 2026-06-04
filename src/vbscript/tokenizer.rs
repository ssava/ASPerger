//! Tokenizer / lexer for the VBScript language. Converts source text
//! into a sequence of `Token` values with associated `TokenType` tags.

use std::iter::Peekable;
use std::str::Chars;

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

#[derive(Debug, Clone)]
pub struct Token {
    pub token_type: TokenType,
    pub value: String,
}

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

        Some(Token {
            token_type: TokenType::EOF,
            value: String::new(),
        })
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

        Token {
            token_type: TokenType::NewLine,
            value,
        }
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

        Token {
            token_type: TokenType::StringLiteral,
            value,
        }
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

        Token { token_type, value }
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

        let token_type = self.get_keyword_type(&value);

        Token { token_type, value }
    }

    fn get_keyword_type(&self, word: &str) -> TokenType {
        if word.eq_ignore_ascii_case("CLASS") {
            return TokenType::Class;
        }
        if word.eq_ignore_ascii_case("FUNCTION") {
            return TokenType::Function;
        }
        if word.eq_ignore_ascii_case("SUB") {
            return TokenType::Sub;
        }
        if word.eq_ignore_ascii_case("DIM") {
            return TokenType::Dim;
        }
        if word.eq_ignore_ascii_case("IF") {
            return TokenType::If;
        }
        if word.eq_ignore_ascii_case("THEN") {
            return TokenType::Then;
        }
        if word.eq_ignore_ascii_case("ELSE") {
            return TokenType::Else;
        }
        if word.eq_ignore_ascii_case("ELSEIF") {
            return TokenType::ElseIf;
        }
        if word.eq_ignore_ascii_case("END") {
            return TokenType::End;
        }
        if word.eq_ignore_ascii_case("FOR") {
            return TokenType::For;
        }
        if word.eq_ignore_ascii_case("NEXT") {
            return TokenType::Next;
        }
        if word.eq_ignore_ascii_case("DO") {
            return TokenType::Do;
        }
        if word.eq_ignore_ascii_case("LOOP") {
            return TokenType::Loop;
        }
        if word.eq_ignore_ascii_case("WHILE") {
            return TokenType::While;
        }
        if word.eq_ignore_ascii_case("WEND") {
            return TokenType::WEnd;
        }
        if word.eq_ignore_ascii_case("SELECT") {
            return TokenType::Select;
        }
        if word.eq_ignore_ascii_case("CASE") {
            return TokenType::Case;
        }
        if word.eq_ignore_ascii_case("WITH") {
            return TokenType::With;
        }
        if word.eq_ignore_ascii_case("SET") {
            return TokenType::Set;
        }
        if word.eq_ignore_ascii_case("NEW") {
            return TokenType::New;
        }
        if word.eq_ignore_ascii_case("TRUE") {
            return TokenType::True;
        }
        if word.eq_ignore_ascii_case("FALSE") {
            return TokenType::False;
        }
        if word.eq_ignore_ascii_case("NOTHING") {
            return TokenType::Nothing;
        }
        if word.eq_ignore_ascii_case("NULL") {
            return TokenType::Null;
        }
        if word.eq_ignore_ascii_case("EMPTY") {
            return TokenType::Empty;
        }
        if word.eq_ignore_ascii_case("AND") {
            return TokenType::And;
        }
        if word.eq_ignore_ascii_case("OR") {
            return TokenType::Or;
        }
        if word.eq_ignore_ascii_case("NOT") {
            return TokenType::Not;
        }
        if word.eq_ignore_ascii_case("MOD") {
            return TokenType::Mod;
        }
        if word.eq_ignore_ascii_case("IS") {
            return TokenType::Is;
        }
        if word.eq_ignore_ascii_case("EQV") {
            return TokenType::Eqv;
        }
        if word.eq_ignore_ascii_case("IMP") {
            return TokenType::Imp;
        }
        if word.eq_ignore_ascii_case("TO") {
            return TokenType::To;
        }
        if word.eq_ignore_ascii_case("STEP") {
            return TokenType::Step;
        }
        if word.eq_ignore_ascii_case("REDIM") {
            return TokenType::ReDim;
        }
        if word.eq_ignore_ascii_case("PRESERVE") {
            return TokenType::Preserve;
        }
        if word.eq_ignore_ascii_case("PROPERTY") {
            return TokenType::Property;
        }
        if word.eq_ignore_ascii_case("GET") {
            return TokenType::Get;
        }
        if word.eq_ignore_ascii_case("LET") {
            return TokenType::Let;
        }
        if word.eq_ignore_ascii_case("PUBLIC") {
            return TokenType::Public;
        }
        if word.eq_ignore_ascii_case("PRIVATE") {
            return TokenType::Private;
        }
        TokenType::Identifier
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

        Token {
            token_type: TokenType::Comment,
            value,
        }
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

        Token {
            token_type: TokenType::DateLiteral,
            value,
        }
    }

    fn tokenize_operator(&mut self) -> Token {
        let mut value = String::new();

        // Get the first character of the operator
        if let Some(&c) = self.input.peek() {
            value.push(c);
            self.advance(); // Consume the operator character

            // Check if the operator is part of a multi-character operator (e.g., >=, <=, ==, etc.)
            match c {
                '+' => {
                    return Token {
                        token_type: TokenType::Plus,
                        value,
                    }
                }
                '-' => {
                    return Token {
                        token_type: TokenType::Minus,
                        value,
                    }
                }
                '*' => {
                    return Token {
                        token_type: TokenType::Multiply,
                        value,
                    }
                }
                '/' => {
                    return Token {
                        token_type: TokenType::Divide,
                        value,
                    }
                }
                '\\' => {
                    return Token {
                        token_type: TokenType::IntDivide,
                        value,
                    }
                }
                '^' => {
                    return Token {
                        token_type: TokenType::Power,
                        value,
                    }
                }
                '&' => {
                    // Check for hex literal: &HFF or &hFF
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
                            return Token {
                                token_type: TokenType::HexLiteral,
                                value,
                            };
                        }
                        // Check for octal literal: &77 or &O77
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
                            return Token {
                                token_type: TokenType::OctLiteral,
                                value,
                            };
                        }
                    }
                    return Token {
                        token_type: TokenType::Concat,
                        value,
                    };
                }
                '=' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance();
                        return Token {
                            token_type: TokenType::Equal,
                            value,
                        };
                    }
                    return Token {
                        token_type: TokenType::Assign,
                        value,
                    };
                }
                '.' => {
                    return Token {
                        token_type: TokenType::Dot,
                        value,
                    }
                }
                ',' => {
                    return Token {
                        token_type: TokenType::Comma,
                        value,
                    }
                }
                ':' => {
                    return Token {
                        token_type: TokenType::Colon,
                        value,
                    }
                }
                '(' => {
                    return Token {
                        token_type: TokenType::LeftParen,
                        value,
                    }
                }
                ')' => {
                    return Token {
                        token_type: TokenType::RightParen,
                        value,
                    }
                }
                '>' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance(); // Consume the '='
                        return Token {
                            token_type: TokenType::GreaterEqual,
                            value,
                        };
                    }
                    return Token {
                        token_type: TokenType::GreaterThan,
                        value,
                    };
                }
                '<' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance(); // Consume the '='
                        return Token {
                            token_type: TokenType::LessEqual,
                            value,
                        };
                    } else if self.input.peek() == Some(&'>') {
                        value.push('>');
                        self.advance(); // Consume the '>'
                        return Token {
                            token_type: TokenType::NotEqual,
                            value,
                        };
                    }
                    return Token {
                        token_type: TokenType::LessThan,
                        value,
                    };
                }
                _ => {
                    // Handle invalid operator
                    self.advance(); // Consume invalid character
                    return Token {
                        token_type: TokenType::Invalid,
                        value,
                    };
                }
            }
        }

        // Default return for unknown operator
        Token {
            token_type: TokenType::EOF,
            value,
        }
    }
}
