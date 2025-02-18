use std::str::Chars;
use std::iter::Peekable;

#[derive(Debug, PartialEq, Clone)]
pub enum TokenType {
    // Keywords
    Class,
    Function,
    Sub,
    Call,
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
    Let,
    New,
    True,
    False,
    Nothing,
    Null,
    Empty,
    Option,
    Explicit,
    Private,
    Public,
    
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
    EOF,
}

#[derive(Debug, Clone)]
pub struct Token {
    pub token_type: TokenType,
    pub value: String,
    pub line: usize,
    pub column: usize,
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
            line: self.current_line,
            column: self.current_column,
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
        let start_column = self.current_column;
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
            line: self.current_line - 1,
            column: start_column,
        }
    }

    fn tokenize_string(&mut self) -> Token {
        let start_column = self.current_column;
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
            line: self.current_line,
            column: start_column,
        }
    }

    fn tokenize_number(&mut self) -> Token {
        let start_column = self.current_column;
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
        
        Token {
            token_type,
            value,
            line: self.current_line,
            column: start_column,
        }
    }

    fn tokenize_identifier(&mut self) -> Token {
        let start_column = self.current_column;
        let mut value = String::new();
        
        while let Some(&c) = self.input.peek() {
            if self.is_identifier_char(c) {
                value.push(c);
                self.advance();
            } else {
                break;
            }
        }
        
        let token_type = self.get_keyword_type(&value.to_uppercase());
        
        Token {
            token_type,
            value,
            line: self.current_line,
            column: start_column,
        }
    }

    fn get_keyword_type(&self, word: &str) -> TokenType {
        match word {
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
            _ => TokenType::Identifier,
        }
    }

    fn is_identifier_start(&self, c: char) -> bool {
        c.is_alphabetic() || c == '_' || c == '['
    }

    fn is_identifier_char(&self, c: char) -> bool {
        c.is_alphanumeric() || c == '_' || c == ']'
    }

    fn is_operator_char(&self, c: char) -> bool {
        matches!(c, '+' | '-' | '*' | '/' | '\\' | '^' | '&' | '=' | '.' | ',' | ':' | '(' | ')' | '>' | '<')
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
        let start_column = self.current_column;
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
            line: self.current_line,
            column: start_column,
        }
    }

    fn tokenize_date(&mut self) -> Token {
        let start_column = self.current_column;
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
            line: self.current_line,
            column: start_column,
        }
    }
    
    fn tokenize_operator(&mut self) -> Token {
        let start_column = self.current_column;
        let mut value = String::new();
        
        // Get the first character of the operator
        if let Some(&c) = self.input.peek() {
            value.push(c);
            self.advance(); // Consume the operator character
            
            // Check if the operator is part of a multi-character operator (e.g., >=, <=, ==, etc.)
            match c {
                '+' => return Token {
                    token_type: TokenType::Plus,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '-' => return Token {
                    token_type: TokenType::Minus,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '*' => return Token {
                    token_type: TokenType::Multiply,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '/' => return Token {
                    token_type: TokenType::Divide,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '\\' => return Token {
                    token_type: TokenType::IntDivide,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '^' => return Token {
                    token_type: TokenType::Power,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '&' => return Token {
                    token_type: TokenType::Concat,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '=' => return Token {
                    token_type: TokenType::Assign,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '.' => return Token {
                    token_type: TokenType::Dot,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                ',' => return Token {
                    token_type: TokenType::Comma,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                ':' => return Token {
                    token_type: TokenType::Colon,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '(' => return Token {
                    token_type: TokenType::LeftParen,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                ')' => return Token {
                    token_type: TokenType::RightParen,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                '>' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance(); // Consume the '='
                        return Token {
                            token_type: TokenType::GreaterEqual,
                            value,
                            line: self.current_line,
                            column: start_column,
                        };
                    }
                    return Token {
                        token_type: TokenType::GreaterThan,
                        value,
                        line: self.current_line,
                        column: start_column,
                    };
                },
                '<' => {
                    if self.input.peek() == Some(&'=') {
                        value.push('=');
                        self.advance(); // Consume the '='
                        return Token {
                            token_type: TokenType::LessEqual,
                            value,
                            line: self.current_line,
                            column: start_column,
                        };
                    } else if self.input.peek() == Some(&'>') {
                        value.push('>');
                        self.advance(); // Consume the '>'
                        return Token {
                            token_type: TokenType::NotEqual,
                            value,
                            line: self.current_line,
                            column: start_column,
                        };
                    }
                    return Token {
                        token_type: TokenType::LessThan,
                        value,
                        line: self.current_line,
                        column: start_column,
                    };
                },
                '=' => return Token {
                    token_type: TokenType::Equal,
                    value,
                    line: self.current_line,
                    column: start_column,
                },
                _ => {
                    // Handle invalid operator
                    self.advance(); // Consume invalid character
                    return Token {
                        token_type: TokenType::EOF, // Represent invalid operator as EOF or handle accordingly
                        value,
                        line: self.current_line,
                        column: start_column,
                    };
                },
            }
        }
    
        // Default return for unknown operator
        Token {
            token_type: TokenType::EOF,
            value,
            line: self.current_line,
            column: start_column,
        }
    }
}