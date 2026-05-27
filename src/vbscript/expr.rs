use super::vbs_error::{VBSError, VBSErrorType};
use super::{ExecutionContext, Token, TokenType, VBValue};

#[derive(Debug, Clone, PartialEq)]
pub enum BinOp {
    Add, Sub, Mul, Div, IntDiv, Pow, Mod, Concat,
    Eq, Ne, Lt, Gt, Le, Ge,
    And, Or, Xor, Eqv, Imp,
    Is,
}

#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOp {
    Neg, Not,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Literal(VBValue),
    Variable(String),
    BinaryOp { left: Box<Expr>, op: BinOp, right: Box<Expr> },
    UnaryOp { op: UnaryOp, expr: Box<Expr> },
    FunctionCall { name: String, args: Vec<Expr> },
    PropertyAccess { object: Box<Expr>, property: String },
    MethodCall { object: Box<Expr>, method: String, args: Vec<Expr> },
}

pub fn parse_expression(tokens: &[Token]) -> Result<Expr, VBSError> {
    let filtered: Vec<&Token> = tokens.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();
    let mut pos = 0;
    let result = parse_binary(&filtered, &mut pos, 0)?;
    Ok(result)
}

fn peek<'a>(tokens: &'a [&'a Token], pos: usize) -> Option<&'a Token> {
    tokens.get(pos).copied()
}

fn advance<'a>(tokens: &'a [&'a Token], pos: &mut usize) -> Option<&'a Token> {
    let t = tokens.get(*pos).copied();
    *pos += 1;
    t
}

fn precedence(token_type: &TokenType) -> (u8, bool) {
    match token_type {
        TokenType::Power => (70, true),
        TokenType::Multiply | TokenType::Divide => (60, false),
        TokenType::IntDivide => (55, false),
        TokenType::Plus | TokenType::Minus => (40, false),
        TokenType::Concat => (35, false),
        TokenType::Assign | TokenType::Equal | TokenType::NotEqual | TokenType::LessThan
        | TokenType::GreaterThan | TokenType::LessEqual | TokenType::GreaterEqual => (30, false),
        TokenType::And => (20, false),
        TokenType::Or => (15, false),
        _ => (0, false),
    }
}

fn token_to_binop(token: &Token) -> Option<BinOp> {
    match token.token_type {
        TokenType::Plus => Some(BinOp::Add),
        TokenType::Minus => Some(BinOp::Sub),
        TokenType::Multiply => Some(BinOp::Mul),
        TokenType::Divide => Some(BinOp::Div),
        TokenType::IntDivide => Some(BinOp::IntDiv),
        TokenType::Power => Some(BinOp::Pow),
        TokenType::Concat => Some(BinOp::Concat),
        TokenType::Equal | TokenType::Assign => Some(BinOp::Eq),
        TokenType::NotEqual => Some(BinOp::Ne),
        TokenType::LessThan => Some(BinOp::Lt),
        TokenType::GreaterThan => Some(BinOp::Gt),
        TokenType::LessEqual => Some(BinOp::Le),
        TokenType::GreaterEqual => Some(BinOp::Ge),
        TokenType::And => Some(BinOp::And),
        TokenType::Or => Some(BinOp::Or),
        TokenType::Mod => Some(BinOp::Mod),
        TokenType::Is => Some(BinOp::Is),
        TokenType::Eqv => Some(BinOp::Eqv),
        TokenType::Imp => Some(BinOp::Imp),
        _ => {
            match token.token_type {
                TokenType::Identifier => {
                    let v = &token.value;
                    if v.eq_ignore_ascii_case("AND") { Some(BinOp::And) }
                    else if v.eq_ignore_ascii_case("OR") { Some(BinOp::Or) }
                    else if v.eq_ignore_ascii_case("MOD") { Some(BinOp::Mod) }
                    else if v.eq_ignore_ascii_case("IS") { Some(BinOp::Is) }
                    else if v.eq_ignore_ascii_case("EQV") { Some(BinOp::Eqv) }
                    else if v.eq_ignore_ascii_case("IMP") { Some(BinOp::Imp) }
                    else if v.eq_ignore_ascii_case("XOR") { Some(BinOp::Xor) }
                    else if v.eq_ignore_ascii_case("NOT") { None }
                    else { None }
                },
                _ => None,
            }
        }
    }
}

fn is_unary_op(token: &Token) -> bool {
    matches!(token.token_type, TokenType::Minus)
        || matches!(token.token_type, TokenType::Plus)
        || (token.token_type == TokenType::Identifier && token.value.eq_ignore_ascii_case("not"))
        || matches!(token.token_type, TokenType::Not)
}

fn parse_primary(tokens: &[&Token], pos: &mut usize) -> Result<Expr, VBSError> {
    let token = advance(tokens, pos).ok_or_else(|| {
        VBSErrorType::SyntaxError.into_error("Unexpected end of expression".to_string())
    })?;

    if is_unary_op(token) {
        let op = match token.token_type {
            TokenType::Minus => UnaryOp::Neg,
            _ => {
                if token.value.eq_ignore_ascii_case("not") || token.token_type == TokenType::Not {
                    UnaryOp::Not
                } else {
                    return parse_primary(tokens, pos);
                }
            }
        };
        let expr = parse_primary(tokens, pos)?;
        return Ok(Expr::UnaryOp { op, expr: Box::new(expr) });
    }

    if matches!(token.token_type, TokenType::LeftParen) {
        let expr = parse_binary(tokens, pos, 0)?;
        let close = advance(tokens, pos).ok_or_else(|| {
            VBSErrorType::SyntaxError.into_error("Expected closing parenthesis".to_string())
        })?;
        if !matches!(close.token_type, TokenType::RightParen) {
            return Err(VBSErrorType::SyntaxError.into_error(
                format!("Expected ')' found '{}'", close.value)
            ));
        }
        return Ok(expr);
    }

    match token.token_type {
        TokenType::IntegerLiteral | TokenType::HexLiteral | TokenType::OctLiteral => {
            let num = parse_numeric_literal(token)?;
            Ok(Expr::Literal(VBValue::Number(num)))
        }
        TokenType::FloatLiteral => {
            let num: f64 = token.value.parse().map_err(|_| {
                VBSErrorType::ValueError.into_error(format!("Invalid float: {}", token.value))
            })?;
            Ok(Expr::Literal(VBValue::Number(num)))
        }
        TokenType::StringLiteral => {
            Ok(Expr::Literal(VBValue::String(token.value.clone())))
        }
        TokenType::True => Ok(Expr::Literal(VBValue::Boolean(true))),
        TokenType::False => Ok(Expr::Literal(VBValue::Boolean(false))),
        TokenType::Null => Ok(Expr::Literal(VBValue::Null)),
        TokenType::Empty => Ok(Expr::Literal(VBValue::Empty)),
        TokenType::Nothing => Ok(Expr::Literal(VBValue::Empty)),
        TokenType::Identifier => {
            let name = token.value.clone();
            match peek(tokens, *pos) {
                Some(next) if next.token_type == TokenType::LeftParen => {
                    advance(tokens, pos);
                    let mut args = Vec::new();
                    loop {
                        if let Some(t) = peek(tokens, *pos) {
                            if t.token_type == TokenType::RightParen {
                                advance(tokens, pos);
                                break;
                            }
                        } else {
                            return Err(VBSErrorType::SyntaxError.into_error(
                                "Unclosed parentheses in function call".to_string()
                            ));
                        }
                        let arg = parse_binary(tokens, pos, 0)?;
                        args.push(arg);
                        match peek(tokens, *pos) {
                            Some(t) if t.token_type == TokenType::Comma => {
                                advance(tokens, pos);
                            }
                            Some(t) if t.token_type == TokenType::RightParen => {
                                advance(tokens, pos);
                                break;
                            }
                            Some(t) => return Err(VBSErrorType::SyntaxError.into_error(
                                format!("Expected ',' or ')' after argument, got '{}'", t.value)
                            )),
                            None => return Err(VBSErrorType::SyntaxError.into_error(
                                "Unclosed parentheses in function call".to_string()
                            )),
                        }
                    }
                    Ok(Expr::FunctionCall { name, args })
                }
                _ => Ok(Expr::Variable(name)),
            }
        }
        _ => Err(VBSErrorType::SyntaxError.into_error(
            format!("Unexpected token in expression: '{}' ({:?})", token.value, token.token_type)
        )),
    }
}

fn parse_numeric_literal(token: &Token) -> Result<f64, VBSError> {
    match token.token_type {
        TokenType::HexLiteral => {
            let hex = token.value.trim_start_matches("&H").trim_start_matches("&h");
            i64::from_str_radix(hex, 16).map(|n| n as f64).map_err(|_| {
                VBSErrorType::ValueError.into_error(format!("Invalid hex: {}", token.value))
            })
        }
        TokenType::OctLiteral => {
            let oct = token.value.trim_start_matches('&');
            i64::from_str_radix(oct, 8).map(|n| n as f64).map_err(|_| {
                VBSErrorType::ValueError.into_error(format!("Invalid octal: {}", token.value))
            })
        }
        TokenType::IntegerLiteral | TokenType::FloatLiteral => {
            token.value.parse::<f64>().map_err(|_| {
                VBSErrorType::ValueError.into_error(format!("Invalid number: {}", token.value))
            })
        }
        _ => Err(VBSErrorType::ValueError.into_error("Not a number".to_string())),
    }
}

fn parse_binary(tokens: &[&Token], pos: &mut usize, min_prec: u8) -> Result<Expr, VBSError> {
    let mut lhs = parse_primary(tokens, pos)?;

    while let Some(token) = peek(tokens, *pos) {
        let prec = if let Some(binop) = token_to_binop(token) {
            match binop {
                BinOp::And => 20,
                BinOp::Or => 15,
                BinOp::Mod => 50,
                BinOp::Is => 30,
                BinOp::Eqv => 8,
                BinOp::Imp => 5,
                BinOp::Xor => 10,
                _ => precedence(&token.token_type).0,
            }
        } else {
            match token.token_type {
                TokenType::RightParen | TokenType::NewLine | TokenType::Then | TokenType::Else
                | TokenType::ElseIf | TokenType::To | TokenType::Step | TokenType::Comma
                | TokenType::EOF => break,
                _ => 0,
            }
        };

        let prec = if prec == 0 && token_to_binop(token).is_some() {
            match token_to_binop(token).unwrap() {
                BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Gt | BinOp::Le | BinOp::Ge => 30,
                BinOp::And => 20,
                BinOp::Or => 15,
                BinOp::Mod => 50,
                BinOp::Is => 30,
                BinOp::Eqv => 8,
                BinOp::Imp => 5,
                BinOp::Xor => 10,
                BinOp::Concat => 35,
                BinOp::Add | BinOp::Sub => 40,
                BinOp::Mul | BinOp::Div => 60,
                BinOp::IntDiv => 55,
                BinOp::Pow => 70,
            }
        } else {
            prec
        };

        if token.token_type == TokenType::To || token.token_type == TokenType::Step {
            break;
        }
        if token.token_type == TokenType::Comma {
            break;
        }

        if prec < min_prec {
            break;
        }

        // Handle property/method access: obj.Prop or obj.Method(args)
        if token.token_type == TokenType::Dot {
            advance(tokens, pos);
            let prop = advance(tokens, pos).ok_or_else(|| {
                VBSErrorType::SyntaxError.into_error("Expected property name after '.'".to_string())
            })?;
            if prop.token_type != TokenType::Identifier {
                return Err(VBSErrorType::SyntaxError.into_error(
                    format!("Expected property name after '.', got '{}'", prop.value)
                ));
            }
            let prop_name = prop.value.clone();

            if let Some(next) = peek(tokens, *pos) {
                if next.token_type == TokenType::LeftParen {
                    advance(tokens, pos);
                    let mut args = Vec::new();
                    loop {
                        if let Some(t) = peek(tokens, *pos) {
                            if t.token_type == TokenType::RightParen {
                                advance(tokens, pos);
                                break;
                            }
                        } else {
                            return Err(VBSErrorType::SyntaxError.into_error(
                                "Unclosed parentheses in method call".to_string()
                            ));
                        }
                        let arg = parse_binary(tokens, pos, 0)?;
                        args.push(arg);
                        match peek(tokens, *pos) {
                            Some(t) if t.token_type == TokenType::Comma => {
                                advance(tokens, pos);
                            }
                            Some(t) if t.token_type == TokenType::RightParen => {
                                advance(tokens, pos);
                                break;
                            }
                            Some(t) => return Err(VBSErrorType::SyntaxError.into_error(
                                format!("Expected ',' or ')' after argument, got '{}'", t.value)
                            )),
                            None => return Err(VBSErrorType::SyntaxError.into_error(
                                "Unclosed parentheses in method call".to_string()
                            )),
                        }
                    }
                    lhs = Expr::MethodCall {
                        object: Box::new(lhs),
                        method: prop_name,
                        args,
                    };
                    continue;
                }
            }

            lhs = Expr::PropertyAccess {
                object: Box::new(lhs),
                property: prop_name,
            };
            continue;
        }

        if let Some(op) = token_to_binop(token) {
            advance(tokens, pos);
            let next_min_prec = prec + 1;
            let rhs = parse_binary(tokens, pos, next_min_prec)?;
            lhs = Expr::BinaryOp {
                left: Box::new(lhs),
                op,
                right: Box::new(rhs),
            };
        } else {
            break;
        }
    }

    Ok(lhs)
}

pub fn evaluate(expr: &Expr, context: &ExecutionContext) -> Result<VBValue, VBSError> {
    match expr {
        Expr::Literal(val) => Ok(val.clone()),
        Expr::Variable(name) => {
            context.get_variable(name).ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error(format!("Variable '{}' is not defined", name))
            })
        }
        Expr::UnaryOp { op, expr } => {
            let val = evaluate(expr, context)?;
            match op {
                UnaryOp::Neg => negate(val),
                UnaryOp::Not => logical_not(val),
            }
        }
        Expr::BinaryOp { left, op, right } => {
            let lv = evaluate(left, context)?;
            let rv = evaluate(right, context)?;
            eval_binary(&lv, op, &rv)
        }
        Expr::FunctionCall { name, args } => {
            let evaluated_args: Result<Vec<VBValue>, VBSError> = args.iter()
                .map(|arg| evaluate(arg, context))
                .collect();
            let evaluated_args = evaluated_args?;

            if let Some(var) = context.get_variable(name) {
                match var {
                    VBValue::Object(obj) => {
                        if evaluated_args.len() == 1 {
                            return obj.indexed_get(&evaluated_args[0]);
                        }
                    }
                    VBValue::Array(ref items) => {
                        if evaluated_args.len() == 1 {
                            let idx = to_number(&evaluated_args[0]) as usize;
                            if idx < items.len() {
                                return Ok(items[idx].clone());
                            }
                            return Err(VBSErrorType::RuntimeError.into_error(
                                format!("Subscript out of range: index {} exceeds array size {}", idx, items.len())
                            ));
                        }
                    }
                    _ => {}
                }
            }
            crate::vbscript::builtins::call_builtin(name, evaluated_args)
        }
        Expr::PropertyAccess { object, property } => {
            let obj_val = evaluate(object, context)?;
            match obj_val {
                VBValue::Object(obj) => obj.get_property(property),
                _ => Err(VBSErrorType::RuntimeError.into_error(
                    format!("Object doesn't support this property or method: '{}'", property)
                )),
            }
        }
        Expr::MethodCall { object, method, args } => {
            let obj_val = evaluate(object, context)?;
            match obj_val {
                VBValue::Object(mut obj) => {
                    let evaluated_args: Result<Vec<VBValue>, VBSError> = args.iter()
                        .map(|arg| evaluate(arg, context))
                        .collect();
                    obj.call_method(method, &evaluated_args?)
                }
                _ => Err(VBSErrorType::RuntimeError.into_error(
                    format!("Object doesn't support this property or method: '{}'", method)
                )),
            }
        }
    }
}

pub(crate) fn to_number(val: &VBValue) -> f64 {
    match val {
        VBValue::Number(n) => *n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty => 0.0,
        VBValue::Array(_) => 0.0,
        VBValue::Object(_) => 0.0,
    }
}

fn to_bool(val: &VBValue) -> bool {
    match val {
        VBValue::Boolean(b) => *b,
        VBValue::Number(n) => *n != 0.0,
        VBValue::String(s) => !s.is_empty(),
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(v) => !v.is_empty(),
        VBValue::Object(_) => true,
    }
}

fn to_string_val(val: &VBValue) -> String {
    match val {
        VBValue::String(s) => s.clone(),
        VBValue::Number(n) => n.to_string(),
        VBValue::Boolean(true) => "True".to_string(),
        VBValue::Boolean(false) => "False".to_string(),
        VBValue::Null => "Null".to_string(),
        VBValue::Empty => "".to_string(),
        VBValue::Array(_) => "Array".to_string(),
        VBValue::Object(_) => "Object".to_string(),
    }
}

fn negate(val: VBValue) -> Result<VBValue, VBSError> {
    if matches!(val, VBValue::Array(_) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    match val {
        VBValue::Number(n) => Ok(VBValue::Number(-n)),
        VBValue::Empty => Ok(VBValue::Number(-0.0)),
        VBValue::Boolean(true) => Ok(VBValue::Number(1.0)),
        VBValue::Boolean(false) => Ok(VBValue::Number(0.0)),
        VBValue::Null => Ok(VBValue::Null),
        VBValue::Array(_) => unreachable!(),
        VBValue::Object(_) => unreachable!(),
        VBValue::String(s) => {
            if let Ok(n) = s.parse::<f64>() {
                Ok(VBValue::Number(-n))
            } else {
                Ok(VBValue::Number(-0.0))
            }
        }
    }
}

fn logical_not(val: VBValue) -> Result<VBValue, VBSError> {
    if matches!(val, VBValue::Array(_) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    Ok(VBValue::Boolean(!to_bool(&val)))
}

fn eval_binary(left: &VBValue, op: &BinOp, right: &VBValue) -> Result<VBValue, VBSError> {
    if matches!(left, VBValue::Array(_) | VBValue::Object(_)) || matches!(right, VBValue::Array(_) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    match op {
        BinOp::Add => {
            match (left, right) {
                (VBValue::String(_), _) | (_, VBValue::String(_))
                    if !matches!(left, VBValue::Number(_)) || !matches!(right, VBValue::Number(_)) =>
                {
                    Ok(VBValue::String(format!("{}{}", to_string_val(left), to_string_val(right))))
                }
                _ => Ok(VBValue::Number(to_number(left) + to_number(right))),
            }
        }
        BinOp::Sub => Ok(VBValue::Number(to_number(left) - to_number(right))),
        BinOp::Mul => Ok(VBValue::Number(to_number(left) * to_number(right))),
        BinOp::Div => {
            let r = to_number(right);
            if r == 0.0 {
                Err(VBSErrorType::RuntimeError.into_error("Division by zero".to_string()))
            } else {
                Ok(VBValue::Number(to_number(left) / r))
            }
        }
        BinOp::IntDiv => {
            let r = to_number(right);
            if r == 0.0 {
                Err(VBSErrorType::RuntimeError.into_error("Division by zero".to_string()))
            } else {
                Ok(VBValue::Number((to_number(left) / r).floor()))
            }
        }
        BinOp::Pow => Ok(VBValue::Number(to_number(left).powf(to_number(right)))),
        BinOp::Mod => Ok(VBValue::Number(to_number(left) % to_number(right))),
        BinOp::Concat => Ok(VBValue::String(format!("{}{}", to_string_val(left), to_string_val(right)))),
        BinOp::Eq => Ok(VBValue::Boolean(values_equal(left, right))),
        BinOp::Ne => Ok(VBValue::Boolean(!values_equal(left, right))),
        BinOp::Lt => Ok(VBValue::Boolean(compare_values(left, right) == std::cmp::Ordering::Less)),
        BinOp::Gt => Ok(VBValue::Boolean(compare_values(left, right) == std::cmp::Ordering::Greater)),
        BinOp::Le => Ok(VBValue::Boolean(compare_values(left, right) != std::cmp::Ordering::Greater)),
        BinOp::Ge => Ok(VBValue::Boolean(compare_values(left, right) != std::cmp::Ordering::Less)),
        BinOp::Is => Ok(VBValue::Boolean(values_equal(left, right))),
        BinOp::And => {
            match (left, right) {
                (VBValue::Boolean(_), VBValue::Boolean(_)) =>
                    Ok(VBValue::Boolean(to_bool(left) && to_bool(right))),
                _ => Ok(VBValue::Number(
                    (to_number(left) as i64 & to_number(right) as i64) as f64
                )),
            }
        }
        BinOp::Or => {
            match (left, right) {
                (VBValue::Boolean(_), VBValue::Boolean(_)) =>
                    Ok(VBValue::Boolean(to_bool(left) || to_bool(right))),
                _ => Ok(VBValue::Number(
                    (to_number(left) as i64 | to_number(right) as i64) as f64
                )),
            }
        }
        BinOp::Xor => {
            match (left, right) {
                (VBValue::Boolean(_), VBValue::Boolean(_)) =>
                    Ok(VBValue::Boolean(to_bool(left) ^ to_bool(right))),
                _ => Ok(VBValue::Number(
                    (to_number(left) as i64 ^ to_number(right) as i64) as f64
                )),
            }
        }
        BinOp::Eqv => {
            let result = if matches!(left, VBValue::Boolean(_)) && matches!(right, VBValue::Boolean(_)) {
                VBValue::Boolean(to_bool(left) == to_bool(right))
            } else {
                VBValue::Number(
                    !(to_number(left) as i64 ^ to_number(right) as i64) as f64
                )
            };
            Ok(result)
        }
        BinOp::Imp => {
            let l = to_bool(left);
            let r = to_bool(right);
            Ok(VBValue::Boolean(!l || r))
        }
    }
}

fn values_equal(left: &VBValue, right: &VBValue) -> bool {
    match (left, right) {
        (VBValue::Number(a), VBValue::Number(b)) => (a - b).abs() < f64::EPSILON,
        (VBValue::String(a), VBValue::String(b)) => a == b,
        (VBValue::Boolean(a), VBValue::Boolean(b)) => a == b,
        (VBValue::Null, VBValue::Null) => true,
        (VBValue::Empty, VBValue::Empty) => true,
        (VBValue::Array(_), _) | (_, VBValue::Array(_)) => false,
        (VBValue::Object(_), _) | (_, VBValue::Object(_)) => false,
        _ => to_string_val(left) == to_string_val(right),
    }
}

fn compare_values(left: &VBValue, right: &VBValue) -> std::cmp::Ordering {
    let ln = to_number(left);
    let rn = to_number(right);
    match ln.partial_cmp(&rn) {
        Some(ordering) => ordering,
        None => std::cmp::Ordering::Equal,
    }
}
