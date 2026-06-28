use super::block_exec::execute_user_defined_function;
use super::block_types::{BlockStatement, CaseClause, ElseIfBlock};
use crate::vbscript::expr::{evaluate, parse_expression, BinOp, Expr};
use crate::vbscript::syntax::{
    ArrayAssignment, Assignment, Const, Dim, MethodCall, OnErrorGoto0, OnErrorResumeNext,
    PropertySet, ReDim, ResponseCookiesSet, ResponseCookiesSetProp, ResponseWrite, VBSyntax,
};
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::{ExecutionContext, Token, TokenType, VBValue};

pub(crate) fn first_non_ws(tokens: &[Token]) -> Option<&Token> {
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

fn tokens_to_string(tokens: &[Token]) -> String {
    tokens
        .iter()
        .map(|t| t.value.as_ref().to_string())
        .collect::<Vec<_>>()
        .join(" ")
}

type DimDeclarations = Vec<(String, Option<Vec<Expr>>)>;

fn parse_dim_statement(tokens: &[Token]) -> Result<DimDeclarations, VBSError> {
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
            i += 1;

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

fn try_parse_response_method(
    tokens: &[Token],
    non_ws: &[&Token],
) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() < 3
        || !non_ws[0].value.eq_ignore_ascii_case("response")
        || non_ws[1].token_type != TokenType::Dot
    {
        return None;
    }
    let method_name = non_ws[2].value.to_string();
    let method_upper = method_name.to_uppercase();
    if method_upper == "END" || method_upper == "CLEAR" || method_upper == "FLUSH" {
        return Some(Ok(Box::new(MethodCall::new(
            "response".to_string(),
            method_name,
            Vec::new(),
        ))));
    }
    if method_upper == "ADDHEADER" || method_upper == "REDIRECT" {
        let arg_tokens: Vec<Token> = tokens
            .iter()
            .skip_while(|t| {
                !(t.token_type == TokenType::Identifier
                    && t.value.eq_ignore_ascii_case(&method_name))
            })
            .skip(1)
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .cloned()
            .collect();
        let args = if arg_tokens.is_empty() {
            Vec::new()
        } else {
            match parse_comma_args(&arg_tokens) {
                Ok(a) => a,
                Err(e) => return Some(Err(e)),
            }
        };
        return Some(Ok(Box::new(MethodCall::new(
            "response".to_string(),
            method_name,
            args,
        ))));
    }
    None
}

fn try_parse_response_write(
    tokens: &[Token],
    non_ws: &[&Token],
) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
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
        if found_write {
            expr_start = i;
            break;
        }
    }
    let expr = if expr_start < tokens.len() {
        parse_expression(&tokens[expr_start..])
    } else {
        Ok(Expr::Literal(VBValue::Empty))
    };
    Some(expr.map(|e| Box::new(ResponseWrite::new(e)) as Box<dyn VBSyntax>))
}

fn try_parse_response_cookies(
    tokens: &[Token],
    non_ws: &[&Token],
) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() < 7
        || !non_ws[0].value.eq_ignore_ascii_case("response")
        || non_ws[1].token_type != TokenType::Dot
        || !non_ws[2].value.eq_ignore_ascii_case("cookies")
        || non_ws[3].token_type != TokenType::LeftParen
    {
        return None;
    }
    let mut i = 0;
    while i < tokens.len()
        && !(tokens[i].token_type == TokenType::Identifier
            && tokens[i].value.eq_ignore_ascii_case("cookies"))
    {
        i += 1;
    }
    i += 1;
    let paren_start = i + 1;
    let mut depth = 1;
    while i < tokens.len() && depth > 0 {
        i += 1;
        if i < tokens.len() {
            if tokens[i].token_type == TokenType::LeftParen {
                depth += 1;
            } else if tokens[i].token_type == TokenType::RightParen {
                depth -= 1;
            }
        }
    }
    let key_expr = match parse_expression(&tokens[paren_start..i]) {
        Ok(e) => e,
        Err(e) => return Some(Err(e)),
    };
    i += 1;
    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }
    if i < tokens.len() && tokens[i].token_type == TokenType::Dot {
        i += 1;
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        if i < tokens.len() && tokens[i].token_type == TokenType::Identifier {
            let property_name = tokens[i].value.to_string();
            i += 1;
            while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
                i += 1;
            }
            if i < tokens.len() && tokens[i].token_type == TokenType::Assign {
                i += 1;
                return Some(parse_expression(&tokens[i..]).map(|v| {
                    Box::new(ResponseCookiesSetProp::new(key_expr, property_name, v))
                        as Box<dyn VBSyntax>
                }));
            }
        }
    }
    if i < tokens.len() && tokens[i].token_type == TokenType::Assign {
        i += 1;
        return Some(
            parse_expression(&tokens[i..])
                .map(|v| Box::new(ResponseCookiesSet::new(key_expr, v)) as Box<dyn VBSyntax>),
        );
    }
    Some(Err(
        VBSErrorType::SyntaxError.into_error("Invalid Response.Cookies syntax".to_string()),
    ))
}

fn try_parse_call_statement(
    tokens: &[Token],
    non_ws: &[&Token],
) -> Option<Result<Box<dyn VBSyntax>, VBSError>> {
    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type != TokenType::Assign
        && non_ws[1].token_type != TokenType::Dot
    {
        let name = non_ws[0].value.to_string();
        if !tokens.iter().any(|t| t.token_type == TokenType::Assign) {
            let name_pos = tokens
                .iter()
                .position(|t| t.token_type == TokenType::Identifier && t.value.as_ref() == name)
                .unwrap_or(0);
            let arg_tokens = &tokens[name_pos + 1..];
            let filtered: Vec<Token> = arg_tokens
                .iter()
                .filter(|t| t.token_type != TokenType::WhiteSpace)
                .cloned()
                .collect();
            let args = if filtered.is_empty() {
                Vec::new()
            } else {
                match parse_comma_args(&filtered) {
                    Ok(a) => a,
                    Err(e) => return Some(Err(e)),
                }
            };
            return Some(Ok(Box::new(CallStatement::new(name, args))));
        }
    }
    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::LeftParen
    {
        let name = non_ws[0].value.to_string();
        let mut i = 0;
        while i < tokens.len() && tokens[i].token_type != TokenType::LeftParen {
            i += 1;
        }
        if i < tokens.len() {
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
            let arg_expr = match parse_expression(&tokens[paren_start..i]) {
                Ok(e) => e,
                Err(e) => return Some(Err(e)),
            };
            return Some(match arg_expr {
                Expr::FunctionCall { ref name, ref args } => {
                    Ok(Box::new(CallStatement::new(name.clone(), args.clone())))
                }
                _ => {
                    let args = if paren_start < i {
                        parse_comma_args(&tokens[paren_start..i])
                    } else {
                        Ok(Vec::new())
                    };
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
        return Ok(Box::new(ArrayAssignment::new(var_name, index_exprs, value_expr)));
    }

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

    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Dot
        && non_ws[1].token_type == TokenType::Identifier
    {
        let method_name = non_ws[1].value.to_string();
        let remaining: Vec<Token> = tokens
            .iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .skip(2)
            .cloned()
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
    let name = no_ws[1].value.to_lowercase();

    let mut params = Vec::new();
    if no_ws.len() > 2 && no_ws[2].token_type == TokenType::LeftParen {
        let mut i = 3;
        while i < no_ws.len() && no_ws[i].token_type != TokenType::RightParen {
            if no_ws[i].token_type == TokenType::Identifier {
                params.push(no_ws[i].value.to_lowercase());
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

fn parse_case_value(tokens: &[Token]) -> Result<Expr, VBSError> {
    if tokens.len() >= 3 && tokens[0].token_type == TokenType::Is {
        let op_token = &tokens[1];
        if let Some(op) = crate::vbscript::expr::token_to_binop(op_token) {
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
        .map(|i| line[i].value.to_lowercase())
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
                    Err(e) => BlockStatement::Unrecognized(e, tokens_to_string(line), line_num),
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
                    then_body: vec![BlockStatement::Unrecognized(
                        VBSErrorType::SyntaxError.into_error(line_text.clone()),
                        line_text,
                        line_num,
                    )],
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
                || (t.token_type == TokenType::Identifier
                    && t.value.eq_ignore_ascii_case("call"))
        })
        .cloned()
        .collect();

    let expr = parse_expression(&rest)?;
    match expr {
        Expr::FunctionCall { name, args } => Ok(Box::new(CallStatement::new(name, args))),
        Expr::MethodCall {
            ref object,
            ref method,
            ref args,
        } => {
            let object_name = expr_to_object_name(object).ok_or_else(|| {
                VBSErrorType::SyntaxError.into_error(
                    "Invalid Call statement: unsupported object expression".to_string(),
                )
            })?;
            Ok(Box::new(MethodCall::new(
                object_name,
                method.clone(),
                args.clone(),
            )))
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
