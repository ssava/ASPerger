//! VBScript block / control-flow parsing and execution.
//! Handles If/ElseIf/Else, For/While/Do loops, Select Case,
//! With/End With, Class/Property definitions, Function/Sub declarations,
//! and top-level statement dispatch.

use super::execution_context::{ClassDefinition, ErrorMode, MethodDef, PropertyDef};
use super::expr::{evaluate, parse_expression, BinOp, Expr};
use super::syntax::{
    ArrayAssignment, Assignment, Const, Dim, MethodCall, OnErrorGoto0, OnErrorResumeNext,
    PropertySet, ReDim, ResponseCookiesSet, ResponseCookiesSetProp, ResponseWrite, VBSyntax,
};
use super::vbs_error::{VBSError, VBSErrorType};
use super::{ExecutionContext, Token, TokenType, VBValue};
use ahash::AHashMap;

/// Parsed VBScript statement, produced by `parse_blocks`.
///
/// Each variant corresponds to a VBScript control-flow or declaration construct.
/// The `line` field (present on all compound variants) stores the source line
/// number used for debugger breakpoint matching.
pub enum BlockStatement {
    Syntax(Box<dyn VBSyntax>, usize),
    Unrecognized(String, usize),
    If {
        line: usize,
        condition: Expr,
        then_body: Vec<BlockStatement>,
        else_if_blocks: Vec<ElseIfBlock>,
        else_body: Option<Vec<BlockStatement>>,
    },
    For {
        line: usize,
        counter: String,
        start: Expr,
        end: Expr,
        step: Option<Expr>,
        body: Vec<BlockStatement>,
    },
    While {
        line: usize,
        condition: Expr,
        body: Vec<BlockStatement>,
    },
    Do {
        line: usize,
        body: Vec<BlockStatement>,
        condition: Option<Expr>,
        is_until: bool,
        is_post_test: bool,
    },
    ForEach {
        line: usize,
        element: String,
        group: Expr,
        body: Vec<BlockStatement>,
    },
    FunctionDef {
        line: usize,
        name: String,
        params: Vec<String>,
        body_lines: Vec<Vec<Token>>,
    },
    SubDef {
        line: usize,
        name: String,
        params: Vec<String>,
        body_lines: Vec<Vec<Token>>,
    },
    SelectCase {
        line: usize,
        expression: Expr,
        cases: Vec<CaseClause>,
        else_body: Option<Vec<BlockStatement>>,
    },
    ClassDef {
        line: usize,
        name: String,
        body_lines: Vec<Vec<Token>>,
    },
    With {
        line: usize,
        object: Expr,
        body: Vec<BlockStatement>,
    },
    ExitFor(usize),
    ExitDo(usize),
    ExitFunction(usize),
    ExitSub(usize),
}

impl Clone for BlockStatement {
    fn clone(&self) -> Self {
        match self {
            BlockStatement::Syntax(s, line) => BlockStatement::Syntax(s.clone_box(), *line),
            BlockStatement::Unrecognized(s, line) => BlockStatement::Unrecognized(s.clone(), *line),
            BlockStatement::If { line, condition, then_body, else_if_blocks, else_body } => {
                BlockStatement::If {
                    line: *line,
                    condition: condition.clone(),
                    then_body: then_body.clone(),
                    else_if_blocks: else_if_blocks.clone(),
                    else_body: else_body.clone(),
                }
            }
            BlockStatement::For { line, counter, start, end, step, body } => {
                BlockStatement::For {
                    line: *line,
                    counter: counter.clone(),
                    start: start.clone(),
                    end: end.clone(),
                    step: step.clone(),
                    body: body.clone(),
                }
            }
            BlockStatement::While { line, condition, body } => {
                BlockStatement::While {
                    line: *line,
                    condition: condition.clone(),
                    body: body.clone(),
                }
            }
            BlockStatement::Do { line, body, condition, is_until, is_post_test } => {
                BlockStatement::Do {
                    line: *line,
                    body: body.clone(),
                    condition: condition.clone(),
                    is_until: *is_until,
                    is_post_test: *is_post_test,
                }
            }
            BlockStatement::ForEach { line, element, group, body } => {
                BlockStatement::ForEach {
                    line: *line,
                    element: element.clone(),
                    group: group.clone(),
                    body: body.clone(),
                }
            }
            BlockStatement::FunctionDef { line, name, params, body_lines } => {
                BlockStatement::FunctionDef {
                    line: *line,
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                }
            }
            BlockStatement::SubDef { line, name, params, body_lines } => {
                BlockStatement::SubDef {
                    line: *line,
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                }
            }
            BlockStatement::SelectCase { line, expression, cases, else_body } => {
                BlockStatement::SelectCase {
                    line: *line,
                    expression: expression.clone(),
                    cases: cases.clone(),
                    else_body: else_body.clone(),
                }
            }
            BlockStatement::ClassDef { line, name, body_lines } => {
                BlockStatement::ClassDef {
                    line: *line,
                    name: name.clone(),
                    body_lines: body_lines.clone(),
                }
            }
            BlockStatement::With { line, object, body } => {
                BlockStatement::With {
                    line: *line,
                    object: object.clone(),
                    body: body.clone(),
                }
            }
            BlockStatement::ExitFor(l) => BlockStatement::ExitFor(*l),
            BlockStatement::ExitDo(l) => BlockStatement::ExitDo(*l),
            BlockStatement::ExitFunction(l) => BlockStatement::ExitFunction(*l),
            BlockStatement::ExitSub(l) => BlockStatement::ExitSub(*l),
        }
    }
}

/// A single `ElseIf condition Then` clause inside an `If` block.
#[derive(Clone)]
pub struct ElseIfBlock {
    pub condition: Expr,
    pub body: Vec<BlockStatement>,
}

/// A single `Case values` clause inside a `Select Case` block.
/// `Case Is operator value` is encoded as `Expr::CaseComparison`.
#[derive(Clone)]
pub struct CaseClause {
    pub values: Vec<Expr>,
    pub body: Vec<BlockStatement>,
}

impl BlockStatement {
    pub fn line(&self) -> usize {
        match self {
            BlockStatement::Syntax(_, l) => *l,
            BlockStatement::Unrecognized(_, l) => *l,
            BlockStatement::If { line: l, .. } => *l,
            BlockStatement::For { line: l, .. } => *l,
            BlockStatement::While { line: l, .. } => *l,
            BlockStatement::Do { line: l, .. } => *l,
            BlockStatement::ForEach { line: l, .. } => *l,
            BlockStatement::FunctionDef { line: l, .. } => *l,
            BlockStatement::SubDef { line: l, .. } => *l,
            BlockStatement::SelectCase { line: l, .. } => *l,
            BlockStatement::ClassDef { line: l, .. } => *l,
            BlockStatement::With { line: l, .. } => *l,
            BlockStatement::ExitFor(l) => *l,
            BlockStatement::ExitDo(l) => *l,
            BlockStatement::ExitFunction(l) => *l,
            BlockStatement::ExitSub(l) => *l,
        }
    }
}

/// A user-defined `Sub` or `Function` parsed from source.
///
/// Function bodies are stored as raw token lines so they can be re-parsed
/// into `BlockStatement`s on each call (VBScript allows redefinition).
/// The cached parsed bodies are stored separately in `ExecutionContext::function_bodies`.
#[derive(Clone)]
pub struct UserDefinedFunction {
    pub name: String,
    pub params: Vec<String>,
    pub body_lines: Vec<Vec<Token>>,
    pub is_function: bool,
}

fn first_non_ws(tokens: &[Token]) -> Option<&Token> {
    tokens
        .iter()
        .find(|t| t.token_type != TokenType::WhiteSpace)
}

fn find_token(tokens: &[Token], target: TokenType) -> Option<usize> {
    tokens.iter().position(|t| t.token_type == target)
}

fn find_keyword_or_type(tokens: &[Token], keyword: &str, token_type: TokenType) -> Option<usize> {
    tokens.iter().position(|t| {
        t.token_type == token_type
            || (t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case(keyword))
    })
}

// ===== Line-level parsing (migrated from VBScriptInterpreter) =====

fn tokens_to_string(tokens: &[Token]) -> String {
    tokens
        .iter()
        .map(|t| t.value.as_ref().to_string())
        .collect::<Vec<_>>()
        .join(" ")
}

#[allow(clippy::type_complexity)]
fn parse_dim_statement(tokens: &[Token]) -> Result<Vec<(String, Option<Vec<Expr>>)>, VBSError> {
    let mut var_names = Vec::new();
    let mut i = 1;

    while i < tokens.len() {
        if tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
            continue;
        }
        if tokens[i].token_type != TokenType::Identifier {
            return Err(VBSErrorType::SyntaxError.into_error(format!(
                "Expected variable name, found: {}",
                tokens[i].value
            )));
        }
        let name = tokens[i].value.to_string();
        i += 1;

        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        let dims = if i < tokens.len() && tokens[i].token_type == TokenType::LeftParen {
            i += 1;
            // Collect tokens until matching RightParen
            let paren_start = i;
            let mut depth = 1;
            while i < tokens.len() && depth > 0 {
                if tokens[i].token_type == TokenType::LeftParen {
                    depth += 1;
                } else if tokens[i].token_type == TokenType::RightParen {
                    depth -= 1;
                }
                if depth > 0 {
                    i += 1;
                }
            }
            if depth != 0 {
                return Err(VBSErrorType::SyntaxError
                    .into_error("Unmatched parentheses in array declaration".to_string()));
            }
            let inner_tokens: Vec<Token> = tokens[paren_start..i]
                .iter()
                .filter(|t| t.token_type != TokenType::WhiteSpace)
                .cloned()
                .collect();
            i += 1; // skip RightParen

            if inner_tokens.is_empty() {
                Some(Vec::new())
            } else {
                let mut dim_exprs = Vec::new();
                let mut expr_start = 0;
                for (j, tok) in inner_tokens.iter().enumerate() {
                    if tok.token_type == TokenType::Comma {
                        if j > expr_start {
                            dim_exprs.push(parse_expression(&inner_tokens[expr_start..j])?);
                        }
                        expr_start = j + 1;
                    }
                }
                if expr_start < inner_tokens.len() {
                    dim_exprs.push(parse_expression(&inner_tokens[expr_start..])?);
                }
                Some(dim_exprs)
            }
        } else {
            None
        };
        var_names.push((name, dims));

        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        if i < tokens.len() && tokens[i].token_type == TokenType::Comma {
            i += 1;
        } else {
            break;
        }
    }

    if var_names.is_empty() {
        return Err(VBSErrorType::SyntaxError
            .into_error("No variable names found in 'Dim' statement".to_string()));
    }
    Ok(var_names)
}

fn parse_const_statement(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let mut const_pairs = Vec::new();
    let mut i = 1;

    while i < tokens.len() {
        if tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
            continue;
        }
        if tokens[i].token_type != TokenType::Identifier {
            return Err(VBSErrorType::SyntaxError.into_error(format!(
                "Expected constant name, found: {}",
                tokens[i].value
            )));
        }
        let name = tokens[i].value.to_string();
        i += 1;

        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        if i >= tokens.len() || tokens[i].token_type != TokenType::Assign {
            return Err(VBSErrorType::SyntaxError
                .into_error("Expected '=' in Const statement".to_string()));
        }
        i += 1;

        let expr = parse_expression(&tokens[i..])?;
        const_pairs.push((name, expr));
        break;
    }

    if const_pairs.is_empty() {
        return Err(VBSErrorType::SyntaxError
            .into_error("No constants defined in 'Const' statement".to_string()));
    }
    Ok(Box::new(Const::new(const_pairs)))
}

fn parse_assignment_statement(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    if tokens.is_empty() {
        return Err(VBSErrorType::SyntaxError.into_error("Empty assignment statement".to_string()));
    }

    let is_set_assignment = tokens[0].token_type == TokenType::Set;
    let mut i = if is_set_assignment { 1 } else { 0 };

    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }

    if i >= tokens.len() || tokens[i].token_type != TokenType::Identifier {
        return Err(VBSErrorType::SyntaxError.into_error(format!(
            "Expected variable name, found: {:?}",
            tokens.get(i)
        )));
    }

    let var_name = tokens[i].value.to_string();
    i += 1;

    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }

    if i >= tokens.len() || tokens[i].token_type != TokenType::Assign {
        return Err(VBSErrorType::SyntaxError
            .into_error(format!("Expected '=', found: {:?}", tokens.get(i))));
    }
    i += 1;

    let expr = parse_expression(&tokens[i..])?;
    Ok(Box::new(Assignment::new(var_name, expr)))
}

fn parse_redim_statement(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let mut i = 1;
    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }

    // Check for Preserve
    let preserve = if i < tokens.len() && tokens[i].token_type == TokenType::Preserve {
        i += 1;
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        true
    } else {
        false
    };

    if i >= tokens.len() || tokens[i].token_type != TokenType::Identifier {
        return Err(
            VBSErrorType::SyntaxError.into_error("Expected variable name after ReDim".to_string())
        );
    }
    let var_name = tokens[i].value.to_string();
    i += 1;

    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }

    if i >= tokens.len() || tokens[i].token_type != TokenType::LeftParen {
        return Err(VBSErrorType::SyntaxError
            .into_error("Expected '(' after variable name in ReDim".to_string()));
    }
    i += 1;

    let paren_start = i;
    let mut depth = 1;
    while i < tokens.len() && depth > 0 {
        if tokens[i].token_type == TokenType::LeftParen {
            depth += 1;
        } else if tokens[i].token_type == TokenType::RightParen {
            depth -= 1;
        }
        if depth > 0 {
            i += 1;
        }
    }
    if depth != 0 {
        return Err(
            VBSErrorType::SyntaxError.into_error("Unmatched parentheses in ReDim".to_string())
        );
    }

    let inner_tokens: Vec<Token> = tokens[paren_start..i]
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();
    let size_exprs = if inner_tokens.is_empty() {
        vec![Expr::Literal(VBValue::Empty)]
    } else {
        let mut exprs = Vec::new();
        let mut expr_start = 0;
        for (j, tok) in inner_tokens.iter().enumerate() {
            if tok.token_type == TokenType::Comma {
                if j > expr_start {
                    exprs.push(parse_expression(&inner_tokens[expr_start..j])?);
                }
                expr_start = j + 1;
            }
        }
        if expr_start < inner_tokens.len() {
            exprs.push(parse_expression(&inner_tokens[expr_start..])?);
        }
        exprs
    };
    Ok(Box::new(ReDim::new(var_name, size_exprs, preserve)))
}

fn find_method_token(tokens: &[Token], method_name: &str) -> Option<usize> {
    let dot_idx = tokens.iter().position(|t| t.token_type == TokenType::Dot)?;
    let start = dot_idx + 1;
    tokens[start..]
        .iter()
        .position(|t| t.token_type == TokenType::Identifier && t.value.as_ref() == method_name)
        .map(|offset| start + offset)
}

fn parse_comma_args(tokens: &[Token]) -> Result<Vec<Expr>, VBSError> {
    if tokens.is_empty() {
        return Ok(Vec::new());
    }
    if !tokens.iter().any(|t| t.token_type == TokenType::Comma) {
        return Ok(vec![parse_expression(tokens)?]);
    }
    let mut args = Vec::new();
    let mut start = 0;
    for (i, tok) in tokens.iter().enumerate() {
        if tok.token_type == TokenType::Comma {
            if i > start {
                let arg_tokens: Vec<Token> = tokens[start..i].to_vec();
                args.push(parse_expression(&arg_tokens)?);
            }
            start = i + 1;
        }
    }
    if start < tokens.len() {
        let arg_tokens: Vec<Token> = tokens[start..].to_vec();
        args.push(parse_expression(&arg_tokens)?);
    }
    Ok(args)
}

/// Disambiguate an ambiguous token sequence into one of several syntax node types.
///
/// Tries patterns in order of specificity:
///  1. Response.Write(expr)
///  2. Response.Cookies("key") = value
///  3. var = expr           (assignment)
///  4. arr(idx) = expr      (array element assignment)
///  5. obj.Property = expr  (property set)
///  6. obj.Method(args)     (method call)
///  7. .Property = value    (With-block property set)
///  8. .Method(args)        (With-block method call)
fn try_parse_response_method(tokens: &[Token], non_ws: &[&Token]) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() < 3 || !non_ws[0].value.eq_ignore_ascii_case("response") || non_ws[1].token_type != TokenType::Dot {
        return None;
    }
    let method_name = non_ws[2].value.to_string();
    let method_upper = method_name.to_uppercase();
    if method_upper == "END" || method_upper == "CLEAR" || method_upper == "FLUSH" {
        return Some(Ok(Box::new(MethodCall::new("response".to_string(), method_name, Vec::new()))));
    }
    if method_upper == "ADDHEADER" || method_upper == "REDIRECT" {
        let arg_tokens: Vec<Token> = tokens.iter()
            .skip_while(|t| !(t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case(&method_name)))
            .skip(1).filter(|t| t.token_type != TokenType::WhiteSpace).cloned().collect();
        let args = if arg_tokens.is_empty() { Vec::new() } else { match parse_comma_args(&arg_tokens) { Ok(a) => a, Err(e) => return Some(Err(e)) } };
        return Some(Ok(Box::new(MethodCall::new("response".to_string(), method_name, args))));
    }
    None
}

fn try_parse_response_write(tokens: &[Token], non_ws: &[&Token]) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() < 3
        || !non_ws[0].value.eq_ignore_ascii_case("response")
        || non_ws[1].token_type != TokenType::Dot
        || !non_ws[2].value.eq_ignore_ascii_case("write")
    {
        return None;
    }
    let mut expr_start = tokens.len();
    let mut found_write = false;
    for (i, tok) in tokens.iter().enumerate() {
        if tok.token_type != TokenType::WhiteSpace && tok.value.eq_ignore_ascii_case("write") {
            found_write = true;
            continue;
        }
        if found_write { expr_start = i; break; }
    }
    let expr = if expr_start < tokens.len() {
        parse_expression(&tokens[expr_start..])
    } else {
        Ok(Expr::Literal(VBValue::Empty))
    };
    Some(expr.map(|e| Box::new(ResponseWrite::new(e)) as Box<dyn VBSyntax>))
}

fn try_parse_response_cookies(tokens: &[Token], non_ws: &[&Token]) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() < 7
        || !non_ws[0].value.eq_ignore_ascii_case("response")
        || non_ws[1].token_type != TokenType::Dot
        || !non_ws[2].value.eq_ignore_ascii_case("cookies")
        || non_ws[3].token_type != TokenType::LeftParen
    {
        return None;
    }
    let mut i = 0;
    while i < tokens.len() && !(tokens[i].token_type == TokenType::Identifier && tokens[i].value.eq_ignore_ascii_case("cookies")) {
        i += 1;
    }
    i += 1;
    let paren_start = i + 1;
    let mut depth = 1;
    while i < tokens.len() && depth > 0 {
        i += 1;
        if i < tokens.len() {
            if tokens[i].token_type == TokenType::LeftParen { depth += 1; }
            else if tokens[i].token_type == TokenType::RightParen { depth -= 1; }
        }
    }
    let key_expr = match parse_expression(&tokens[paren_start..i]) {
        Ok(e) => e, Err(e) => return Some(Err(e)),
    };
    i += 1;
    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace { i += 1; }
    if i < tokens.len() && tokens[i].token_type == TokenType::Dot {
        i += 1;
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace { i += 1; }
        if i < tokens.len() && tokens[i].token_type == TokenType::Identifier {
            let property_name = tokens[i].value.to_string();
            i += 1;
            while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace { i += 1; }
            if i < tokens.len() && tokens[i].token_type == TokenType::Assign {
                i += 1;
                return Some(parse_expression(&tokens[i..]).map(|v| Box::new(ResponseCookiesSetProp::new(key_expr, property_name, v)) as Box<dyn VBSyntax>));
            }
        }
    }
    if i < tokens.len() && tokens[i].token_type == TokenType::Assign {
        i += 1;
        return Some(parse_expression(&tokens[i..]).map(|v| Box::new(ResponseCookiesSet::new(key_expr, v)) as Box<dyn VBSyntax>));
    }
    Some(Err(VBSErrorType::SyntaxError.into_error("Invalid Response.Cookies syntax".to_string())))
}

fn try_parse_call_statement(tokens: &[Token], non_ws: &[&Token]) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() >= 2 && non_ws[0].token_type == TokenType::Identifier && non_ws[1].token_type != TokenType::Assign && non_ws[1].token_type != TokenType::Dot {
        let name = non_ws[0].value.to_string();
        if !tokens.iter().any(|t| t.token_type == TokenType::Assign) {
            let name_pos = tokens.iter().position(|t| t.token_type == TokenType::Identifier && t.value.as_ref() == name).unwrap_or(0);
            let arg_tokens = &tokens[name_pos + 1..];
            let filtered: Vec<Token> = arg_tokens.iter().filter(|t| t.token_type != TokenType::WhiteSpace).cloned().collect();
            let args = if filtered.is_empty() { Vec::new() } else { match parse_comma_args(&filtered) { Ok(a) => a, Err(e) => return Some(Err(e)) } };
            return Some(Ok(Box::new(CallStatement::new(name, args))));
        }
    }
    if non_ws.len() >= 2 && non_ws[0].token_type == TokenType::Identifier && non_ws[1].token_type == TokenType::LeftParen {
        let name = non_ws[0].value.to_string();
        let mut i = 0;
        while i < tokens.len() && tokens[i].token_type != TokenType::LeftParen { i += 1; }
        if i < tokens.len() {
            i += 1;
            let paren_start = i;
            let mut depth = 1;
            while i < tokens.len() && depth > 0 {
                if tokens[i].token_type == TokenType::LeftParen { depth += 1; }
                else if tokens[i].token_type == TokenType::RightParen { depth -= 1; }
                if depth > 0 { i += 1; }
            }
            let arg_expr = match parse_expression(&tokens[paren_start..i]) {
                Ok(e) => e, Err(e) => return Some(Err(e)),
            };
            return Some(match arg_expr {
                Expr::FunctionCall { ref name, ref args } => Ok(Box::new(CallStatement::new(name.clone(), args.clone()))),
                _ => {
                    let args = if paren_start < i { parse_comma_args(&tokens[paren_start..i]) } else { Ok(Vec::new()) };
                    args.map(|a| Box::new(CallStatement::new(name, a)) as Box<dyn VBSyntax>)
                }
            });
        }
    }
    None
}

fn parse_expression_or_assignment(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let non_ws: Vec<&Token> = tokens
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if let Some(result) = try_parse_response_method(tokens, &non_ws) {
        return result;
    }
    if let Some(result) = try_parse_response_write(tokens, &non_ws) {
        return result;
    }
    if let Some(result) = try_parse_response_cookies(tokens, &non_ws) {
        return result;
    }

    // var = expr (bare assignment, no Set keyword)
    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::Assign
    {
        let var_name = non_ws[0].value.to_string();
        let assign_idx = tokens
            .iter()
            .position(|t| t.token_type == TokenType::Assign)
            .unwrap();
        let expr = parse_expression(&tokens[assign_idx + 1..])?;
        return Ok(Box::new(Assignment::new(var_name, expr)));
    }

    // 4. arr(idx) = expr (array element assignment)
    if non_ws.len() >= 4
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::LeftParen
    {
        let var_name = non_ws[0].value.to_string();
        let mut i = 0;
        while i < tokens.len() && tokens[i].token_type != TokenType::LeftParen {
            i += 1;
        }
        if i < tokens.len() && tokens[i].token_type == TokenType::LeftParen {
            i += 1;
        }
        let paren_start = i;
        let mut depth = 1;
        while i < tokens.len() && depth > 0 {
            if tokens[i].token_type == TokenType::LeftParen {
                depth += 1;
            } else if tokens[i].token_type == TokenType::RightParen {
                depth -= 1;
            }
            if depth > 0 {
                i += 1;
            }
        }
        if depth != 0 {
            return Err(VBSErrorType::SyntaxError
                .into_error("Unmatched parentheses in array assignment".to_string()));
        }
        let inner_tokens: Vec<Token> = tokens[paren_start..i]
            .iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .cloned()
            .collect();
        let index_exprs = if inner_tokens.is_empty() {
            vec![Expr::Literal(VBValue::Empty)]
        } else {
            let mut exprs = Vec::new();
            let mut expr_start = 0;
            for (j, tok) in inner_tokens.iter().enumerate() {
                if tok.token_type == TokenType::Comma {
                    if j > expr_start {
                        exprs.push(parse_expression(&inner_tokens[expr_start..j])?);
                    }
                    expr_start = j + 1;
                }
            }
            if expr_start < inner_tokens.len() {
                exprs.push(parse_expression(&inner_tokens[expr_start..])?);
            }
            exprs
        };
        i += 1;
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        if i >= tokens.len() || tokens[i].token_type != TokenType::Assign {
            return Err(
                VBSErrorType::SyntaxError.into_error("Expected '=' after array index".to_string())
            );
        }
        i += 1;
        let value_expr = parse_expression(&tokens[i..])?;
        return Ok(Box::new(ArrayAssignment::new(
            var_name, index_exprs, value_expr,
        )));
    }

    // obj.Property = expr (property set)
    if non_ws.len() >= 4
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::Dot
        && non_ws[2].token_type == TokenType::Identifier
        && non_ws[3].token_type == TokenType::Assign
    {
        let object_name = non_ws[0].value.to_string();
        let property_name = non_ws[2].value.to_string();
        let assign_idx = tokens
            .iter()
            .position(|t| t.token_type == TokenType::Assign)
            .unwrap();
        let value_expr = parse_expression(&tokens[assign_idx + 1..])?;
        return Ok(Box::new(PropertySet::new(
            object_name,
            property_name,
            value_expr,
        )));
    }

    // obj.Method arg1, arg2, ... (method call)
    if non_ws.len() >= 3
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::Dot
        && non_ws[2].token_type == TokenType::Identifier
    {
        let object_name = non_ws[0].value.to_string();
        let method_name = non_ws[2].value.to_string();

        let args = if let Some(mi) = find_method_token(tokens, &method_name) {
            let arg_tokens = &tokens[mi + 1..];
            parse_comma_args(arg_tokens)?
        } else {
            Vec::new()
        };

        return Ok(Box::new(MethodCall::new(object_name, method_name, args)));
    }

    // With-block: .Property = value (property set)
    if non_ws.len() >= 3
        && non_ws[0].token_type == TokenType::Dot
        && non_ws[1].token_type == TokenType::Identifier
        && non_ws[2].token_type == TokenType::Assign
    {
        let property_name = non_ws[1].value.to_string();
        let assign_idx = tokens
            .iter()
            .position(|t| t.token_type == TokenType::Assign)
            .unwrap();
        let value_expr = parse_expression(&tokens[assign_idx + 1..])?;
        return Ok(Box::new(PropertySet::new(
            "__with_obj__".to_string(),
            property_name,
            value_expr,
        )));
    }

    // With-block: .Method arg1, arg2, ... (method call)
    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Dot
        && non_ws[1].token_type == TokenType::Identifier
    {
        let method_name = non_ws[1].value.to_string();
        let remaining: Vec<Token> = tokens
            .iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .skip(2).cloned()
            .collect();
        let args = if remaining.is_empty() {
            Vec::new()
        } else {
            parse_comma_args(&remaining)?
        };
        return Ok(Box::new(MethodCall::new(
            "__with_obj__".to_string(),
            method_name,
            args,
        )));
    }

    if let Some(result) = try_parse_call_statement(tokens, &non_ws) {
        return result;
    }

    Err(VBSErrorType::NotImplementedError.into_error(format!(
        "Unrecognized command: {}",
        tokens_to_string(tokens)
    )))
}

fn parse_line_into_syntax(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let first_token = tokens
        .iter()
        .find(|t| t.token_type != TokenType::WhiteSpace)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("Empty statement".to_string()))?;

    match first_token.token_type {
        TokenType::Dim => {
            let var_names = parse_dim_statement(tokens)?;
            Ok(Box::new(Dim::new(var_names)))
        }
        TokenType::Set => parse_assignment_statement(tokens),
        TokenType::ReDim => parse_redim_statement(tokens),
        TokenType::Const => parse_const_statement(tokens),
        TokenType::Identifier if first_token.value.eq_ignore_ascii_case("call") => {
            parse_call_statement(tokens)
        }
        TokenType::Identifier if first_token.value.eq_ignore_ascii_case("on") => {
            parse_on_error_statement(tokens)
        }
        TokenType::Identifier if first_token.value.eq_ignore_ascii_case("exit") => {
            parse_exit_line(tokens)
        }
        TokenType::Identifier if first_token.value.eq_ignore_ascii_case("randomize") => {
            Ok(Box::new(CallStatement::new("Randomize".to_string(), Vec::new())))
        }
        _ => parse_expression_or_assignment(tokens),
    }
}

// ===== Token-to-Expr helpers =====

// ===== Function/Sub parsing =====

fn parse_function_def(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let no_ws: Vec<&Token> = line
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if no_ws.len() < 2 {
        return Err(VBSErrorType::SyntaxError.into_error("Expected Function/Sub name".to_string()));
    }

    let is_function = no_ws[0].token_type == TokenType::Function;
    let name = no_ws[1].value.to_string();

    let mut params = Vec::new();
    if no_ws.len() > 2 && no_ws[2].token_type == TokenType::LeftParen {
        let mut i = 3;
        while i < no_ws.len() && no_ws[i].token_type != TokenType::RightParen {
            if no_ws[i].token_type == TokenType::Identifier {
                params.push(no_ws[i].value.to_string());
            }
            i += 1;
        }
    }

    let body_start = *pos;

    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error(format!(
                "{} without End {}",
                if is_function { "Function" } else { "Sub" },
                if is_function { "Function" } else { "Sub" }
            )));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::End => {
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if (is_function
                        && (s.value.eq_ignore_ascii_case("function")
                            || s.token_type == TokenType::Function))
                        || (!is_function
                            && (s.value.eq_ignore_ascii_case("sub")
                                || s.token_type == TokenType::Sub))
                    {
                        let body_lines: Vec<Vec<Token>> = lines[body_start..*pos].to_vec();
                        *pos += 1;
                        if is_function {
                            return Ok(BlockStatement::FunctionDef {
                                line: line_num,
                                name,
                                params,
                                body_lines,
                            });
                        } else {
                            return Ok(BlockStatement::SubDef {
                                line: line_num,
                                name,
                                params,
                                body_lines,
                            });
                        }
                    }
                }
                *pos += 1;
            }
            _ => {
                *pos += 1;
            }
        }
    }
}

// ===== Select Case parsing =====

/// Parse a single Case value expression, handling `Is <op> expr` syntax.
fn parse_case_value(tokens: &[Token]) -> Result<Expr, VBSError> {
    if tokens.len() >= 3
        && tokens[0].token_type == TokenType::Is
    {
        let op_token = &tokens[1];
        if let Some(op) = super::expr::token_to_binop(op_token) {
            if matches!(
                op,
                BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Gt | BinOp::Le | BinOp::Ge
            ) {
                let rhs = parse_expression(&tokens[2..])?;
                return Ok(Expr::CaseComparison {
                    op,
                    rhs: Box::new(rhs),
                });
            }
        }
    }
    parse_expression(tokens)
}

fn parse_select_case_block(
    lines: &[Vec<Token>],
    pos: &mut usize,
) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let case_idx = find_keyword_or_type(line, "case", TokenType::Case)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("Select without Case".to_string()))?;

    let expression = parse_expr_from_slice(line, case_idx + 1)?;

    let mut cases: Vec<CaseClause> = Vec::new();
    let mut else_body: Option<Vec<BlockStatement>> = None;
    let mut in_else = false;

    loop {
        if *pos >= lines.len() {
            return Err(
                VBSErrorType::SyntaxError.into_error("Select without End Select".to_string())
            );
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::Case => {
                let no_ws: Vec<&Token> = next_line
                    .iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .collect();

                if no_ws.len() > 1 && no_ws[1].value.eq_ignore_ascii_case("else") {
                    *pos += 1;
                    in_else = true;
                    else_body = Some(Vec::new());
                } else {
                    let case_token_pos = next_line
                        .iter()
                        .position(|t| t.token_type == TokenType::Case)
                        .unwrap();
                    let after_case: Vec<Token> = next_line[case_token_pos + 1..]
                        .iter()
                        .filter(|t| t.token_type != TokenType::WhiteSpace)
                        .cloned()
                        .collect();

                    let mut values = Vec::new();
                    if !after_case.is_empty() {
                        let mut val_start = 0;
                        for (j, tok) in after_case.iter().enumerate() {
                            if tok.token_type == TokenType::Comma {
                                if j > val_start {
                                    values.push(parse_case_value(&after_case[val_start..j])?);
                                }
                                val_start = j + 1;
                            }
                        }
                        if val_start < after_case.len() {
                            values.push(parse_case_value(&after_case[val_start..])?);
                        }
                    }

                    *pos += 1;
                    in_else = false;
                    cases.push(CaseClause {
                        values,
                        body: Vec::new(),
                    });
                }
            }
            Some(t) if t.token_type == TokenType::End => {
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if s.value.eq_ignore_ascii_case("select") || s.token_type == TokenType::Select {
                        *pos += 1;
                        break;
                    }
                }
                return Err(VBSErrorType::SyntaxError.into_error(format!(
                    "Expected End Select, got: {}",
                    tokens_to_string(next_line)
                )));
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                if in_else {
                    if let Some(ref mut eb) = else_body {
                        eb.extend(sub_blocks);
                    }
                } else if let Some(last) = cases.last_mut() {
                    last.body.extend(sub_blocks);
                }
            }
        }
    }

    Ok(BlockStatement::SelectCase {
        line: line_num,
        expression,
        cases,
        else_body,
    })
}

fn parse_class_def(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let name_idx = line
        .iter()
        .skip_while(|t| t.token_type == TokenType::WhiteSpace)
        .position(|t| t.token_type != TokenType::Class && t.token_type == TokenType::Identifier)
        .and_then(|_idx| {
            let mut count = 0;
            for (i, t) in line.iter().enumerate() {
                if t.token_type == TokenType::WhiteSpace {
                    continue;
                }
                if count == 1 {
                    return Some(i);
                }
                count += 1;
            }
            None
        });

    let class_name = name_idx
        .map(|i| line[i].value.to_string())
        .unwrap_or_else(|| "".to_string());

    let mut body_lines: Vec<Vec<Token>> = Vec::new();
    let mut depth = 0;

    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("Class without End Class".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::End => {
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if s.value.eq_ignore_ascii_case("class") || s.token_type == TokenType::Class {
                        if depth == 0 {
                            *pos += 1;
                            break;
                        }
                        depth -= 1;
                    } else if s.value.eq_ignore_ascii_case("property")
                        || s.token_type == TokenType::Property
                    {
                        if depth == 0 {
                            return Err(VBSErrorType::SyntaxError
                                .into_error("End Property without matching Property".to_string()));
                        }
                        depth -= 1;
                    }
                }
                body_lines.push(next_line.clone());
                *pos += 1;
            }
            Some(t) if t.token_type == TokenType::Property => {
                depth += 1;
                body_lines.push(next_line.clone());
                *pos += 1;
            }
            Some(t) if t.token_type == TokenType::Public || t.token_type == TokenType::Private => {
                // Check if followed by Property or Class — need depth tracking
                // for End Property / End Class matching
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if s.token_type == TokenType::Property || s.token_type == TokenType::Class {
                        depth += 1;
                    }
                }
                body_lines.push(next_line.clone());
                *pos += 1;
            }
            _ => {
                body_lines.push(next_line.clone());
                *pos += 1;
            }
        }
    }

    Ok(BlockStatement::ClassDef {
        line: line_num,
        name: class_name,
        body_lines,
    })
}

// ===== Block parsing =====

/// Parse a sequence of tokenized lines into a tree of `BlockStatement` nodes.
pub fn parse_blocks(lines: &[Vec<Token>]) -> Result<Vec<BlockStatement>, VBSError> {
    let mut pos = 0;
    parse_blocks_inner(lines, &mut pos)
}

fn parse_blocks_inner(
    lines: &[Vec<Token>],
    pos: &mut usize,
) -> Result<Vec<BlockStatement>, VBSError> {
    let mut blocks = Vec::new();

    while *pos < lines.len() {
        let line = &lines[*pos];
        let first = first_non_ws(line);
        let line_num = *pos + 1;

        match first {
            Some(t) if t.token_type == TokenType::If => {
                blocks.push(parse_if_block(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::For => {
                blocks.push(parse_for_block(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::While => {
                blocks.push(parse_while_block(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::Do => {
                blocks.push(parse_do_block(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::Function || t.token_type == TokenType::Sub => {
                blocks.push(parse_function_def(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::Select => {
                blocks.push(parse_select_case_block(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::Class => {
                blocks.push(parse_class_def(lines, pos)?);
            }
            Some(t) if t.token_type == TokenType::With => {
                blocks.push(parse_with_block(lines, pos)?);
            }
            Some(t)
                if t.token_type == TokenType::End
                    || t.token_type == TokenType::Next
                    || t.token_type == TokenType::WEnd
                    || t.token_type == TokenType::Loop
                    || t.token_type == TokenType::ElseIf
                    || t.token_type == TokenType::Else
                    || t.token_type == TokenType::Case
                    || t.token_type == TokenType::Property
                    || t.token_type == TokenType::Public
                    || t.token_type == TokenType::Private =>
            {
                break;
            }
            Some(t)
                if t.token_type == TokenType::Identifier
                    && t.value.eq_ignore_ascii_case("exit") =>
            {
                blocks.push(parse_exit_statement(lines, pos)?);
            }
            _ => {
                // Skip comment lines
                if line.iter().any(|t| {
                    t.token_type == TokenType::Comment
                        || (t.token_type == TokenType::Identifier
                            && t.value.eq_ignore_ascii_case("rem"))
                }) {
                    *pos += 1;
                    continue;
                }
                blocks.push(match parse_line_into_syntax(line) {
                    Ok(syntax) => BlockStatement::Syntax(syntax, line_num),
                    Err(_) => BlockStatement::Unrecognized(tokens_to_string(line), line_num),
                });
                *pos += 1;
            }
        }
    }

    Ok(blocks)
}

fn parse_expr_from_range(tokens: &[Token], start: usize, end: usize) -> Result<Expr, VBSError> {
    parse_expression(&tokens[start..end])
}

fn parse_expr_from_slice(tokens: &[Token], start: usize) -> Result<Expr, VBSError> {
    parse_expression(&tokens[start..])
}

fn parse_if_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let then_idx = find_keyword_or_type(line, "then", TokenType::Then)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("If without Then".to_string()))?;

    let condition = parse_expr_from_range(line, 1, then_idx)?;

    let after_then: Vec<&Token> = line[then_idx + 1..]
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if !after_then.is_empty() {
        let inline_tokens: Vec<Token> = line[then_idx + 1..].to_vec();
        let line_text = tokens_to_string(&inline_tokens);
        let syntax = match parse_line_into_syntax(&inline_tokens) {
            Ok(s) => s,
            Err(_) => {
                return Ok(BlockStatement::If {
                    line: line_num,
                    condition,
                    then_body: vec![BlockStatement::Unrecognized(line_text, line_num)],
                    else_if_blocks: Vec::new(),
                    else_body: None,
                })
            }
        };
        let then_body = vec![BlockStatement::Syntax(syntax, line_num)];
        return Ok(BlockStatement::If {
            line: line_num,
            condition,
            then_body,
            else_if_blocks: Vec::new(),
            else_body: None,
        });
    }

    enum Section {
        Then,
        ElseIf(usize),
        Else,
    }

    let mut then_body: Vec<BlockStatement> = Vec::new();
    let mut else_if_blocks: Vec<ElseIfBlock> = Vec::new();
    let mut else_body: Option<Vec<BlockStatement>> = None;
    let mut section = Section::Then;

    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("If without End If".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::End => {
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if s.value.eq_ignore_ascii_case("if") || s.token_type == TokenType::If {
                        *pos += 1;
                        break;
                    }
                }
                let line_text = tokens_to_string(next_line);
                let line = *pos + 1;
                let syntax = parse_line_into_syntax(next_line)
                    .unwrap_or_else(|_| Box::new(create_error_syntax(line_text.clone())));
                match &section {
                    Section::Then => then_body.push(BlockStatement::Syntax(syntax, line)),
                    Section::ElseIf(idx) => else_if_blocks[*idx]
                        .body
                        .push(BlockStatement::Syntax(syntax, line)),
                    Section::Else => {
                        if let Some(ref mut eb) = else_body {
                            eb.push(BlockStatement::Syntax(syntax, line));
                        }
                    }
                }
                *pos += 1;
            }
            Some(t) if t.token_type == TokenType::ElseIf => {
                let elseif_line = &lines[*pos];
                *pos += 1;

                let then_idx = find_keyword_or_type(elseif_line, "then", TokenType::Then)
                    .ok_or_else(|| {
                        VBSErrorType::SyntaxError.into_error("ElseIf without Then".to_string())
                    })?;

                let elseif_cond = parse_expr_from_range(elseif_line, 1, then_idx)?;

                let after_then: Vec<&Token> = elseif_line[then_idx + 1..]
                    .iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .collect();

                if !after_then.is_empty() {
                    let inline_tokens: Vec<Token> = elseif_line[then_idx + 1..].to_vec();
                    let line_text = tokens_to_string(&inline_tokens);
                    let syntax = match parse_line_into_syntax(&inline_tokens) {
                        Ok(s) => s,
                        Err(_) => Box::new(create_error_syntax(line_text)),
                    };
                    let inline_body = vec![BlockStatement::Syntax(syntax, *pos)];
                    else_if_blocks.push(ElseIfBlock {
                        condition: elseif_cond,
                        body: inline_body,
                    });
                    else_body.get_or_insert_with(Vec::new);
                    section = Section::Else;
                } else {
                    else_if_blocks.push(ElseIfBlock {
                        condition: elseif_cond,
                        body: Vec::new(),
                    });
                    section = Section::ElseIf(else_if_blocks.len() - 1);
                }
            }
            Some(t) if t.token_type == TokenType::Else => {
                *pos += 1;
                else_body = Some(Vec::new());
                section = Section::Else;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                match &section {
                    Section::Then => then_body.extend(sub_blocks),
                    Section::ElseIf(idx) => else_if_blocks[*idx].body.extend(sub_blocks),
                    Section::Else => {
                        if let Some(ref mut eb) = else_body {
                            eb.extend(sub_blocks);
                        }
                    }
                }
            }
        }
    }

    Ok(BlockStatement::If {
        line: line_num,
        condition,
        then_body,
        else_if_blocks,
        else_body,
    })
}

fn parse_for_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let for_line_no_ws: Vec<&Token> = line
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if for_line_no_ws.len() < 5 {
        return Err(VBSErrorType::SyntaxError.into_error("Invalid For statement".to_string()));
    }

    let counter = for_line_no_ws[1].value.to_string();

    if counter.eq_ignore_ascii_case("each") {
        return parse_for_each_block(line, line_num, pos, lines, &for_line_no_ws);
    }

    let assign_idx = find_token(line, TokenType::Assign)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("For without =".to_string()))?;

    let to_idx = find_keyword_or_type(line, "to", TokenType::To)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("For without To".to_string()))?;

    let step_idx = find_keyword_or_type(line, "step", TokenType::Step);

    let start = parse_expr_from_range(line, assign_idx + 1, to_idx)?;

    let end = if let Some(si) = step_idx {
        parse_expr_from_range(line, to_idx + 1, si)?
    } else {
        parse_expr_from_slice(line, to_idx + 1)?
    };

    let step: Option<Expr> = step_idx
        .map(|si| parse_expr_from_slice(line, si + 1))
        .transpose()?;

    let mut body = Vec::new();
    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("For without Next".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::Next => {
                *pos += 1;
                break;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    Ok(BlockStatement::For {
        line: line_num,
        counter,
        start,
        end,
        step,
        body,
    })
}

fn parse_for_each_block(
    line: &[Token],
    line_num: usize,
    pos: &mut usize,
    lines: &[Vec<Token>],
    for_line_no_ws: &[&Token],
) -> Result<BlockStatement, VBSError> {
    if for_line_no_ws.len() < 5 {
        return Err(VBSErrorType::SyntaxError.into_error("Invalid For Each statement".to_string()));
    }

    let element = for_line_no_ws[2].value.to_string();

    let in_idx = line
        .iter()
        .position(|t| t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("in"))
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("For Each without In".to_string()))?;

    let group = parse_expr_from_slice(line, in_idx + 1)?;

    let mut body = Vec::new();
    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("For Each without Next".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::Next => {
                *pos += 1;
                break;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    Ok(BlockStatement::ForEach {
        line: line_num,
        element,
        group,
        body,
    })
}

fn parse_while_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let condition = parse_expr_from_slice(line, 1)?;

    let mut body = Vec::new();
    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("While without Wend".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::WEnd => {
                *pos += 1;
                break;
            }
            Some(t)
                if t.token_type == TokenType::Identifier
                    && t.value.eq_ignore_ascii_case("wend") =>
            {
                *pos += 1;
                break;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    Ok(BlockStatement::While {
        line: line_num,
        condition,
        body,
    })
}

fn parse_do_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let do_line_no_ws: Vec<&Token> = line
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    let mut is_pre_test = false;
    let mut is_until = false;
    let mut pre_condition: Option<Expr> = None;

    if do_line_no_ws.len() > 1 {
        let second = &do_line_no_ws[1];
        if second.value.eq_ignore_ascii_case("while") || second.token_type == TokenType::While {
            is_pre_test = true;
            is_until = false;
            let while_idx = find_keyword_or_type(line, "while", TokenType::While).unwrap();
            pre_condition = Some(parse_expr_from_slice(line, while_idx + 1)?);
        } else if second.value.eq_ignore_ascii_case("until") {
            is_pre_test = true;
            is_until = true;
            let until_idx = line
                .iter()
                .position(|t| {
                    t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("until")
                })
                .unwrap();
            pre_condition = Some(parse_expr_from_slice(line, until_idx + 1)?);
        }
    }

    let mut body = Vec::new();
    let mut post_condition: Option<Expr> = None;
    let mut is_post_until = false;

    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("Do without Loop".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::Loop => {
                let loop_line = &lines[*pos];
                *pos += 1;

                let loop_no_ws: Vec<&Token> = loop_line
                    .iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .collect();
                if loop_no_ws.len() > 1 {
                    let second = &loop_no_ws[1];
                    if second.value.eq_ignore_ascii_case("while")
                        || second.token_type == TokenType::While
                    {
                        is_post_until = false;
                        let while_idx =
                            find_keyword_or_type(loop_line, "while", TokenType::While).unwrap();
                        post_condition = Some(parse_expr_from_slice(loop_line, while_idx + 1)?);
                    } else if second.value.eq_ignore_ascii_case("until") {
                        is_post_until = true;
                        let until_idx = loop_line
                            .iter()
                            .position(|t| {
                                t.token_type == TokenType::Identifier
                                    && t.value.eq_ignore_ascii_case("until")
                            })
                            .unwrap();
                        post_condition = Some(parse_expr_from_slice(loop_line, until_idx + 1)?);
                    }
                }
                break;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    if is_pre_test {
        Ok(BlockStatement::Do {
            line: line_num,
            body,
            condition: pre_condition,
            is_until,
            is_post_test: false,
        })
    } else {
        Ok(BlockStatement::Do {
            line: line_num,
            body,
            condition: post_condition,
            is_until: is_post_until,
            is_post_test: true,
        })
    }
}

fn parse_with_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let with_idx = line
        .iter()
        .position(|t| t.token_type == TokenType::With)
        .unwrap_or(0);
    let object = parse_expr_from_slice(line, with_idx + 1)?;

    let mut body = Vec::new();
    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("With without End With".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::End => {
                let second = next_line
                    .iter()
                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                    .skip(1)
                    .find(|t| t.token_type != TokenType::WhiteSpace);
                if let Some(s) = second {
                    if s.value.eq_ignore_ascii_case("with") || s.token_type == TokenType::With {
                        *pos += 1;
                        break;
                    }
                }
                // Not End With; skip line to avoid infinite loop
                *pos += 1;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    Ok(BlockStatement::With {
        line: line_num,
        object,
        body,
    })
}

fn parse_exit_statement(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    let line_num = *pos + 1;
    *pos += 1;

    let no_ws: Vec<&Token> = line
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    let exit_type = no_ws
        .iter()
        .skip(1)
        .find(|t| t.token_type != TokenType::WhiteSpace)
        .map(|t| t.value.as_ref())
        .unwrap_or("");

    match exit_type.to_uppercase().as_str() {
        "FOR" => Ok(BlockStatement::ExitFor(line_num)),
        "DO" => Ok(BlockStatement::ExitDo(line_num)),
        "FUNCTION" => Ok(BlockStatement::ExitFunction(line_num)),
        "SUB" => Ok(BlockStatement::ExitSub(line_num)),
        _ => Err(VBSErrorType::SyntaxError
            .into_error(format!("Invalid Exit statement: Exit {}", exit_type))),
    }
}

// ===== Execution =====

/// Fast-path: evaluate a simple `Variable <op> Literal` condition using native f64 ops.
/// Returns `None` if the condition doesn't match the simple pattern.
fn try_fast_condition(expr: &Expr, context: &mut ExecutionContext) -> Option<Result<bool, VBSError>> {
    use std::cmp::Ordering;
    let (var_name, lit_val, op, swap) = match expr {
        Expr::BinaryOp { left, op, right } => {
            match (left.as_ref(), right.as_ref()) {
                (Expr::Variable(v), Expr::Literal(VBValue::Number(n))) => {
                    (v.clone(), *n, op, false)
                }
                (Expr::Literal(VBValue::Number(n)), Expr::Variable(v)) => {
                    (v.clone(), *n, op, true)
                }
                _ => return None,
            }
        }
        _ => return None,
    };
    let var_val = match context.get_variable(&var_name) {
        Some(VBValue::Number(n)) => *n,
        Some(VBValue::Empty | VBValue::Null) => 0.0,
        _ => return None,
    };
    let cmp = if var_val < lit_val {
        Ordering::Less
    } else if var_val > lit_val {
        Ordering::Greater
    } else {
        Ordering::Equal
    };
    let result = if swap {
        match op {
            BinOp::Eq => cmp == Ordering::Equal,
            BinOp::Ne => cmp != Ordering::Equal,
            BinOp::Le => cmp != Ordering::Less,
            BinOp::Ge => cmp != Ordering::Greater,
            BinOp::Lt => cmp == Ordering::Greater,
            BinOp::Gt => cmp == Ordering::Less,
            _ => return None,
        }
    } else {
        match op {
            BinOp::Eq => cmp == Ordering::Equal,
            BinOp::Ne => cmp != Ordering::Equal,
            BinOp::Le => cmp != Ordering::Greater,
            BinOp::Ge => cmp != Ordering::Less,
            BinOp::Lt => cmp == Ordering::Less,
            BinOp::Gt => cmp == Ordering::Greater,
            _ => return None,
        }
    };
    Some(Ok(result))
}

fn evaluate_condition(expr: &Expr, context: &mut ExecutionContext) -> Result<bool, VBSError> {
    // Try fast path for simple variable <op> literal comparisons
    if let Some(result) = try_fast_condition(expr, context) {
        return result;
    }
    let val = evaluate(expr, context)?;
    if matches!(val, VBValue::Array(..) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    Ok(match val {
        VBValue::Boolean(b) => b,
        VBValue::Number(n) => n != 0.0,
        VBValue::String(s) => !s.is_empty(),
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(..) => unreachable!(),
        VBValue::Object(_) => unreachable!(),
    })
}

fn evaluate_numeric(expr: &Expr, context: &mut ExecutionContext) -> Result<f64, VBSError> {
    let val = evaluate(expr, context)?;
    if matches!(val, VBValue::Array(..) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    Ok(match val {
        VBValue::Number(n) => n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty => 0.0,
        VBValue::Array(..) => unreachable!(),
        VBValue::Object(_) => unreachable!(),
    })
}

#[derive(Clone)]
struct ErrorSyntax {
    message: String,
}

impl VBSyntax for ErrorSyntax {
    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
    fn execute(&self, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::NotImplementedError
            .into_error(format!("Unrecognized command: {}", self.message)))
    }
}

fn create_error_syntax(message: String) -> ErrorSyntax {
    ErrorSyntax { message }
}

#[derive(Clone)]
struct ExitSyntax {
    exit_type: VBSErrorType,
    label: String,
}

impl VBSyntax for ExitSyntax {
    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }

    fn execute(&self, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(self.exit_type.into_error(self.label.clone()))
    }
}

fn parse_exit_line(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let non_ws: Vec<&Token> = tokens
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();
    let exit_kind = non_ws.get(1).map(|t| t.value.as_ref()).unwrap_or("");
    match exit_kind.to_uppercase().as_str() {
        "FOR" => Ok(Box::new(ExitSyntax {
            exit_type: VBSErrorType::ExitFor,
            label: "Exit For".to_string(),
        })),
        "DO" => Ok(Box::new(ExitSyntax {
            exit_type: VBSErrorType::ExitDo,
            label: "Exit Do".to_string(),
        })),
        "FUNCTION" => Ok(Box::new(ExitSyntax {
            exit_type: VBSErrorType::ExitFunction,
            label: "Exit Function".to_string(),
        })),
        "SUB" => Ok(Box::new(ExitSyntax {
            exit_type: VBSErrorType::ExitSub,
            label: "Exit Sub".to_string(),
        })),
        _ => Err(VBSErrorType::SyntaxError
            .into_error(format!("Invalid Exit statement: Exit {}", exit_kind))),
    }
}

#[derive(Clone)]
pub(crate) struct CallStatement {
    name: String,
    args: Vec<Expr>,
}

impl CallStatement {
    pub(crate) fn new(name: String, args: Vec<Expr>) -> Self {
        CallStatement { name, args }
    }
}

impl VBSyntax for CallStatement {
    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }

    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let args: Result<Vec<VBValue>, VBSError> =
            self.args.iter().map(|arg| evaluate(arg, context)).collect();
        let args = args?;

        if let Some(func) = context.get_function(&self.name).cloned() {
            execute_user_defined_function(&func, &args, context)?;
            return Ok(());
        }

        match crate::vbscript::builtins::call_builtin(&self.name, args) {
            Ok(_) => Ok(()),
            Err(e) => Err(e),
        }
    }
}

fn expr_to_object_name(expr: &Expr) -> Option<String> {
    match expr {
        Expr::Variable(name) => Some(name.clone()),
        Expr::PropertyAccess { object, property } => {
            let base = expr_to_object_name(object)?;
            Some(format!("{}.{}", base, property))
        }
        _ => None,
    }
}

fn parse_call_statement(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let rest: Vec<Token> = tokens
        .iter()
        .skip_while(|t| {
            t.token_type == TokenType::WhiteSpace
                || (t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("call"))
        })
        .cloned()
        .collect();

    let expr = parse_expression(&rest)?;
    match expr {
        Expr::FunctionCall { name, args } => Ok(Box::new(CallStatement::new(name, args))),
        Expr::MethodCall { ref object, ref method, ref args } => {
            let object_name = expr_to_object_name(object).ok_or_else(|| {
                VBSErrorType::SyntaxError
                    .into_error("Invalid Call statement: unsupported object expression".to_string())
            })?;
            Ok(Box::new(MethodCall::new(object_name, method.clone(), args.clone())))
        }
        _ => Err(VBSErrorType::SyntaxError
            .into_error("Invalid Call statement: expected function call".to_string())),
    }
}

fn parse_on_error_statement(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let non_ws: Vec<&Token> = tokens
        .iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();
    if non_ws.len() >= 4
        && non_ws[0].value.eq_ignore_ascii_case("on")
        && non_ws[1].value.eq_ignore_ascii_case("error")
        && non_ws[2].value.eq_ignore_ascii_case("resume")
        && non_ws[3].value.eq_ignore_ascii_case("next")
    {
        return Ok(Box::new(OnErrorResumeNext));
    }
    if non_ws.len() >= 4
        && non_ws[0].value.eq_ignore_ascii_case("on")
        && non_ws[1].value.eq_ignore_ascii_case("error")
        && non_ws[2].value.eq_ignore_ascii_case("goto")
        && non_ws[3].token_type == TokenType::IntegerLiteral
        && non_ws[3].value.as_ref() == "0"
    {
        return Ok(Box::new(OnErrorGoto0));
    }
    Err(VBSErrorType::SyntaxError.into_error(format!(
        "Invalid On Error statement: {}",
        tokens_to_string(tokens)
    )))
}

pub(crate) fn execute_user_defined_function(
    func: &UserDefinedFunction,
    args: &[VBValue],
    context: &mut ExecutionContext,
) -> Result<VBValue, VBSError> {
    // Push stack frame for debugger
    if let Some(ref debugger) = context.debugger {
        use super::execution_context::CIString;
        let vars: AHashMap<CIString, VBValue> = func
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| (CIString::new(p.clone()), args.get(i).cloned().unwrap_or(VBValue::Empty)))
            .collect();
        debugger.push_frame(&func.name, &context.script_path, 0, vars);
    }

    for (i, param) in func.params.iter().enumerate() {
        let val = args.get(i).cloned().unwrap_or(VBValue::Empty);
        context.set_variable(param, val);
    }

    if func.is_function {
        context.set_variable(&func.name, VBValue::Empty);
    }

    // Save and reset code_start_line for function bodies so
    // breakpoints inside functions use logical (function-relative) line numbers
    let saved_code_start_line = context.code_start_line;
    context.code_start_line = 0;

    let body_blocks = match context.get_function_body(&func.name) {
        Some(cached) => cached.clone(),
        None => {
            let blocks = parse_blocks(&func.body_lines)?;
            let name = func.name.clone();
            context.set_function_body(&name, blocks.clone());
            blocks
        }
    };
    match execute_blocks(&body_blocks, context) {
        Ok(()) => {}
        Err(e) if e.is_exit_function() || e.is_exit_sub() => {}
        Err(e) => {
            context.code_start_line = saved_code_start_line;
            return Err(e);
        }
    }

    context.code_start_line = saved_code_start_line;

    // Pop stack frame for debugger
    if let Some(ref debugger) = context.debugger {
        debugger.pop_frame();
    }

    if func.is_function {
        Ok(context
            .get_variable(&func.name)
            .cloned()
            .unwrap_or(VBValue::Empty))
    } else {
        Ok(VBValue::Empty)
    }
}

fn extract_properties_from_class_body(
    body_lines: &[Vec<Token>],
) -> Result<AHashMap<String, PropertyDef>, VBSError> {
    let mut properties: AHashMap<String, PropertyDef> = AHashMap::new();
    let mut i = 0;

    while i < body_lines.len() {
        let line = &body_lines[i];
        let no_ws: Vec<&Token> = line
            .iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .collect();

        if no_ws.is_empty()
            || (no_ws[0].token_type == TokenType::Public
                || no_ws[0].token_type == TokenType::Private)
        {
            let property_idx = no_ws
                .iter()
                .position(|t| t.token_type == TokenType::Property);
            if let Some(p_idx) = property_idx {
                let get_let_set = no_ws.get(p_idx + 1);
                let name_tok = no_ws.get(p_idx + 2);
                let is_get = get_let_set
                    .map(|t| t.token_type == TokenType::Get || t.value.eq_ignore_ascii_case("get"))
                    .unwrap_or(false);
                let is_let = get_let_set
                    .map(|t| t.token_type == TokenType::Let || t.value.eq_ignore_ascii_case("let"))
                    .unwrap_or(false);
                if is_get || is_let {
                    let name_tok = match name_tok {
                        Some(t) if t.token_type == TokenType::Identifier => t,
                        _ => {
                            i += 1;
                            continue;
                        }
                    };
                    let prop_name = name_tok.value.to_string();
                    i += 1;

                    let mut param = None;
                    if is_let && no_ws.len() > p_idx + 3 {
                        let paren_open = no_ws.get(p_idx + 3);
                        if paren_open
                            .map(|t| t.token_type == TokenType::LeftParen)
                            .unwrap_or(false)
                        {
                            if let Some(param_tok) = no_ws.get(p_idx + 4) {
                                if param_tok.token_type == TokenType::Identifier {
                                    param = Some(param_tok.value.to_string());
                                }
                            }
                        }
                    }

                    let mut body: Vec<Vec<Token>> = Vec::new();
                    loop {
                        if i >= body_lines.len() {
                            return Err(VBSErrorType::SyntaxError
                                .into_error("Property without End Property".to_string()));
                        }
                        let bline = &body_lines[i];
                        let first = first_non_ws(bline);
                        if let Some(f) = first {
                            if f.token_type == TokenType::End {
                                let second = bline
                                    .iter()
                                    .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                                    .skip(1)
                                    .find(|t| t.token_type != TokenType::WhiteSpace);
                                if let Some(s) = second {
                                    if s.value.eq_ignore_ascii_case("property")
                                        || s.token_type == TokenType::Property
                                    {
                                        i += 1;
                                        break;
                                    }
                                }
                            }
                        }
                        body.push(bline.clone());
                        i += 1;
                    }

                    let entry = properties
                        .entry(prop_name.to_uppercase())
                        .or_insert(PropertyDef {
                            name: prop_name.clone(),
                            get_body: None,
                            let_body: None,
                            let_param: None,
                        });

                    if is_get {
                        entry.get_body = Some(body);
                    } else if is_let {
                        entry.let_body = Some(body);
                        entry.let_param = param;
                    }
                    continue;
                }
            }
        }
        i += 1;
    }

    Ok(properties)
}

fn extract_methods_from_class_body(body_lines: &[Vec<Token>]) -> AHashMap<String, MethodDef> {
    let mut methods: AHashMap<String, MethodDef> = AHashMap::new();
    let mut i = 0;

    while i < body_lines.len() {
        let line = &body_lines[i];
        let no_ws: Vec<&Token> = line
            .iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .collect();

        if no_ws.is_empty() {
            i += 1;
            continue;
        }

        let start_idx = if no_ws[0].token_type == TokenType::Public
            || no_ws[0].token_type == TokenType::Private
        {
            if no_ws.len() < 2 {
                i += 1;
                continue;
            }
            1
        } else {
            0
        };

        let is_func = no_ws[start_idx].token_type == TokenType::Function
            || no_ws[start_idx].value.eq_ignore_ascii_case("function");
        let is_sub = no_ws[start_idx].token_type == TokenType::Sub
            || no_ws[start_idx].value.eq_ignore_ascii_case("sub");

        if !is_func && !is_sub {
            i += 1;
            continue;
        }

        let name_idx = start_idx + 1;
        if name_idx >= no_ws.len() || no_ws[name_idx].token_type != TokenType::Identifier {
            i += 1;
            continue;
        }
        let method_name = no_ws[name_idx].value.to_string();

        let mut params = Vec::new();
        if no_ws.len() > name_idx + 1 && no_ws[name_idx + 1].token_type == TokenType::LeftParen {
            let mut p = name_idx + 2;
            while p < no_ws.len() && no_ws[p].token_type != TokenType::RightParen {
                if no_ws[p].token_type == TokenType::Identifier {
                    params.push(no_ws[p].value.to_string());
                }
                p += 1;
            }
        }

        i += 1;
        let mut body: Vec<Vec<Token>> = Vec::new();

        loop {
            if i >= body_lines.len() {
                break;
            }
            let bline = &body_lines[i];
            let first = first_non_ws(bline);
            if let Some(f) = first {
                if f.token_type == TokenType::End {
                    let second = bline
                        .iter()
                        .skip_while(|t| t.token_type == TokenType::WhiteSpace)
                        .skip(1)
                        .find(|t| t.token_type != TokenType::WhiteSpace);
                    if let Some(s) = second {
                        if (is_func
                            && (s.value.eq_ignore_ascii_case("function")
                                || s.token_type == TokenType::Function))
                            || (is_sub
                                && (s.value.eq_ignore_ascii_case("sub")
                                    || s.token_type == TokenType::Sub))
                        {
                            i += 1;
                            break;
                        }
                    }
                }
            }
            body.push(bline.clone());
            i += 1;
        }

        let key = method_name.to_uppercase();
        if !methods.contains_key(&key) {
            methods.insert(
                key,
                MethodDef {
                    name: method_name,
                    params,
                    body_lines: body,
                    is_function: is_func,
                },
            );
        }
    }

    methods
}

/// Execute a slice of `BlockStatement` nodes in order, handling control flow
/// (Exit For/Do/Function/Sub) and debugger hooks.
pub fn execute_blocks(
    blocks: &[BlockStatement],
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    tracing::trace!(block_count = blocks.len(), "Executing VB blocks");
    for block in blocks {
        tracing::trace!(line = block.line(), "Executing block");
        // Check if Response.End or Response.Redirect was called
        if context.response.ended {
            break;
        }

        // Debugger hook: check breakpoints and stepping
        if let Some(ref debugger) = context.debugger {
            let frame_depth = debugger.current_frame_depth();
            let file_line = if context.code_start_line > 0 {
                block.line() + context.code_start_line - 1
            } else {
                block.line()
            };
            debugger.check(
                &context.script_path,
                file_line,
                frame_depth,
                Some(context.variables()),
            )?;
        }

        match block {
            // --- Function/Sub definitions: register for later use ---
            BlockStatement::FunctionDef {
                name,
                params,
                body_lines,
                ..
            } => {
                if context.get_function_body(name).is_none() {
                    let bodies = parse_blocks(body_lines)?;
                    context.set_function_body(name, bodies);
                }
                context.define_function(UserDefinedFunction {
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                    is_function: true,
                });
            }
            BlockStatement::SubDef {
                name,
                params,
                body_lines,
                ..
            } => {
                if context.get_function_body(name).is_none() {
                    let bodies = parse_blocks(body_lines)?;
                    context.set_function_body(name, bodies);
                }
                context.define_function(UserDefinedFunction {
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                    is_function: false,
                });
            }
            // --- Syntax nodes (assignment, method call, dim, etc.) ---
            BlockStatement::Syntax(syntax, _line) => {
                // On Error Resume Next: record error and continue
                let result = syntax.execute(context);
                if let Err(e) = result {
                    if *context.get_error_mode() == ErrorMode::ResumeNext {
                        context.set_err(e);
                    } else {
                        return Err(e);
                    }
                }
            }
            // --- Parse failure: unrecognised statement ---
            BlockStatement::Unrecognized(line_text, _line) => {
                return Err(VBSErrorType::NotImplementedError
                    .into_error(format!("Unrecognized command: {}", line_text)));
            }
            // --- If / ElseIf / Else ---
            BlockStatement::If {
                condition,
                then_body,
                else_if_blocks,
                else_body,
                ..
            } => {
                if evaluate_condition(condition, context)? {
                    execute_blocks(then_body, context)?;
                } else {
                    let mut handled = false;
                    for elseif in else_if_blocks {
                        if evaluate_condition(&elseif.condition, context)? {
                            execute_blocks(&elseif.body, context)?;
                            handled = true;
                            break;
                        }
                    }
                    if !handled {
                        if let Some(body) = else_body {
                            execute_blocks(body, context)?;
                        }
                    }
                }
            }
            // --- For (numeric counter step) ---
            BlockStatement::For {
                counter,
                start,
                end,
                step,
                body,
                ..
            } => {
                let start_val = evaluate_numeric(start, context)?;
                let end_val = evaluate_numeric(end, context)?;
                let step_val = step
                    .as_ref()
                    .map(|s| evaluate_numeric(s, context))
                    .unwrap_or(Ok(1.0))?;

                let mut i = start_val;
                if step_val > 0.0 {
                    while i <= end_val {
                        context.set_variable(counter, VBValue::Number(i));
                        match execute_blocks(body, context) {
                            Ok(()) => {}
                            Err(e) if e.is_exit_for() => break,
                            Err(e) => return Err(e),
                        }
                        i += step_val;
                    }
                } else if step_val < 0.0 {
                    while i >= end_val {
                        context.set_variable(counter, VBValue::Number(i));
                        match execute_blocks(body, context) {
                            Ok(()) => {}
                            Err(e) if e.is_exit_for() => break,
                            Err(e) => return Err(e),
                        }
                        i += step_val;
                    }
                }
                context.set_variable(counter, VBValue::Number(i));
            }
            // --- For Each (array/collection iteration) ---
            BlockStatement::ForEach {
                element,
                group,
                body,
                ..
            } => {
                let group_val = evaluate(group, context)?;
                match group_val {
                    VBValue::Array(ref items, _) => {
                        for item in items.iter() {
                            context.set_variable(element, item.clone());
                            match execute_blocks(body, context) {
                                Ok(()) => {}
                                Err(e) if e.is_exit_for() => break,
                                Err(e) => return Err(e),
                            }
                        }
                    }
                    _ => {
                        return Err(VBSErrorType::RuntimeError.into_error(
                            "Object doesn't support this property or method".to_string(),
                        ));
                    }
                }
            }
            // --- While condition loop ---
            BlockStatement::While {
                condition, body, ..
            } => {
                while evaluate_condition(condition, context)? {
                    match execute_blocks(body, context) {
                        Ok(()) => {}
                        Err(e) if e.is_exit_do() => break,
                        Err(e) => return Err(e),
                    }
                }
            }
            // --- Do [While/Until] [condition] / Loop [While/Until] [condition] ---
            BlockStatement::Do {
                body,
                condition,
                is_until,
                is_post_test,
                ..
            } => {
                if *is_post_test {
                    loop {
                        match execute_blocks(body, context) {
                            Ok(()) => {}
                            Err(e) if e.is_exit_do() => break,
                            Err(e) => return Err(e),
                        }
                        if let Some(cond) = condition {
                            let result = evaluate_condition(cond, context)?;
                            if *is_until {
                                if result {
                                    break;
                                }
                            } else if !result {
                                break;
                            }
                        } else {
                            break;
                        }
                    }
                } else {
                    loop {
                        if let Some(cond) = condition {
                            let result = evaluate_condition(cond, context)?;
                            if *is_until {
                                if result {
                                    break;
                                }
                            } else if !result {
                                break;
                            }
                        }
                        match execute_blocks(body, context) {
                            Ok(()) => {}
                            Err(e) if e.is_exit_do() => break,
                            Err(e) => return Err(e),
                        }
                    }
                }
            }
            // --- Select Case ---
            BlockStatement::SelectCase {
                expression,
                cases,
                else_body,
                ..
            } => {
                let expr_val = evaluate(expression, context)?;
                context.select_value = Some(expr_val.clone());
                let mut matched = false;

                for case in cases {
                    for val_expr in &case.values {
                        let case_val = evaluate(val_expr, context)?;
                        let is_match = match val_expr {
                            Expr::CaseComparison { .. } => {
                                matches!(case_val, VBValue::Boolean(true))
                            }
                            _ => val_equal(&expr_val, &case_val),
                        };
                        if is_match {
                            execute_blocks(&case.body, context)?;
                            matched = true;
                            break;
                        }
                    }
                    if matched {
                        break;
                    }
                }
                context.select_value = None;

                if !matched {
                    if let Some(body) = else_body {
                        execute_blocks(body, context)?;
                    }
                }
            }
            // --- Class definition: extract properties and register ---
            BlockStatement::ClassDef {
                name, body_lines, ..
            } => {
                if name.is_empty() {
                    return Err(
                        VBSErrorType::SyntaxError.into_error("Class name is empty".to_string())
                    );
                }
                if let Ok(properties) = extract_properties_from_class_body(body_lines) {
                    let methods = extract_methods_from_class_body(body_lines);
                    let class_def = ClassDefinition {
                        name: name.clone(),
                        properties,
                        methods,
                    };
                    context.define_class(class_def);
                }
            }
            // --- With block: swap scope with-object, restore after body ---
            BlockStatement::With { object, body, .. } => {
                let obj_val = evaluate(object, context)?;
                let prev_with = context.with_object.replace(obj_val);
                let result = execute_blocks(body, context);
                context.with_object = prev_with;
                result?
            }
            // --- Control-flow sentinels propagated as errors ---
            BlockStatement::ExitFor(_) => {
                return Err(VBSErrorType::ExitFor.into_error("Exit For".to_string()));
            }
            BlockStatement::ExitDo(_) => {
                return Err(VBSErrorType::ExitDo.into_error("Exit Do".to_string()));
            }
            BlockStatement::ExitFunction(_) => {
                return Err(VBSErrorType::ExitFunction.into_error("Exit Function".to_string()));
            }
            BlockStatement::ExitSub(_) => {
                return Err(VBSErrorType::ExitSub.into_error("Exit Sub".to_string()));
            }
        }
    }
    Ok(())
}

fn val_equal(a: &VBValue, b: &VBValue) -> bool {
    match (a, b) {
        (VBValue::Number(an), VBValue::Number(bn)) => an == bn,
        (VBValue::String(as_), VBValue::String(bs)) => as_ == bs,
        (VBValue::Boolean(ab), VBValue::Boolean(bb)) => ab == bb,
        (VBValue::Empty, VBValue::Empty) => true,
        (VBValue::Null, VBValue::Null) => true,
        (VBValue::Number(an), VBValue::String(bs)) => {
            bs.parse::<f64>().map(|bn| an == &bn).unwrap_or(false)
        }
        (VBValue::String(as_), VBValue::Number(bn)) => {
            as_.parse::<f64>().map(|an| &an == bn).unwrap_or(false)
        }
        _ => false,
    }
}
