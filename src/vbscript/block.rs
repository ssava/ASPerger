use super::vbs_error::{VBSError, VBSErrorType};
use super::expr::{evaluate, parse_expression, Expr};
use super::syntax::{ArrayAssignment, Assignment, Dim, MethodCall, ReDim, ResponseWrite, VBSyntax};
use super::{ExecutionContext, Token, TokenType, VBValue};

pub enum BlockStatement {
    Syntax(Box<dyn VBSyntax>),
    Unrecognized(String),
    If {
        condition: Expr,
        then_body: Vec<BlockStatement>,
        else_if_blocks: Vec<ElseIfBlock>,
        else_body: Option<Vec<BlockStatement>>,
    },
    For {
        counter: String,
        start: Expr,
        end: Expr,
        step: Option<Expr>,
        body: Vec<BlockStatement>,
    },
    While {
        condition: Expr,
        body: Vec<BlockStatement>,
    },
    Do {
        body: Vec<BlockStatement>,
        condition: Option<Expr>,
        is_until: bool,
        is_post_test: bool,
    },
    ForEach {
        element: String,
        group: Expr,
        body: Vec<BlockStatement>,
    },
}

pub struct ElseIfBlock {
    pub condition: Expr,
    pub body: Vec<BlockStatement>,
}

fn first_non_ws<'a>(tokens: &'a [Token]) -> Option<&'a Token> {
    tokens.iter().find(|t| t.token_type != TokenType::WhiteSpace)
}

fn find_token(tokens: &[Token], target: TokenType) -> Option<usize> {
    tokens.iter().position(|t| t.token_type == target)
}

fn find_keyword_or_type(tokens: &[Token], keyword: &str, token_type: TokenType) -> Option<usize> {
    tokens.iter().position(|t| {
        t.token_type == token_type ||
        (t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case(keyword))
    })
}

// ===== Line-level parsing (migrated from VBScriptInterpreter) =====

fn tokens_to_string(tokens: &[Token]) -> String {
    tokens.iter().map(|t| t.value.clone()).collect::<Vec<String>>().join(" ")
}

fn parse_dim_statement(tokens: &[Token]) -> Result<Vec<(String, bool)>, VBSError> {
    let mut var_names = Vec::new();
    let mut i = 1;

    while i < tokens.len() {
        if tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
            continue;
        }
        if tokens[i].token_type != TokenType::Identifier {
            return Err(VBSErrorType::SyntaxError.into_error(
                format!("Expected variable name, found: {}", tokens[i].value)
            ));
        }
        let name = tokens[i].value.clone();
        i += 1;

        // Check for array declaration: arr()
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        let is_array = if i < tokens.len() && tokens[i].token_type == TokenType::LeftParen {
            i += 1;
            while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
                i += 1;
            }
            if i >= tokens.len() || tokens[i].token_type != TokenType::RightParen {
                return Err(VBSErrorType::SyntaxError.into_error(
                    "Expected ')' after '(' in array declaration".to_string()
                ));
            }
            i += 1;
            true
        } else {
            false
        };
        var_names.push((name, is_array));

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
        return Err(VBSErrorType::SyntaxError.into_error("No variable names found in 'Dim' statement".to_string()));
    }
    Ok(var_names)
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
        return Err(VBSErrorType::SyntaxError.into_error(
            "Expected variable name after ReDim".to_string()
        ));
    }
    let var_name = tokens[i].value.clone();
    i += 1;

    while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
        i += 1;
    }

    if i >= tokens.len() || tokens[i].token_type != TokenType::LeftParen {
        return Err(VBSErrorType::SyntaxError.into_error(
            "Expected '(' after variable name in ReDim".to_string()
        ));
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
        return Err(VBSErrorType::SyntaxError.into_error(
            "Unmatched parentheses in ReDim".to_string()
        ));
    }

    let size_expr = parse_expression(&tokens[paren_start..i])?;
    Ok(Box::new(ReDim::new(var_name, size_expr, preserve)))
}

fn find_method_token(tokens: &[Token], method_name: &str) -> Option<usize> {
    let dot_idx = tokens.iter().position(|t| t.token_type == TokenType::Dot)?;
    let start = dot_idx + 1;
    tokens[start..].iter().position(|t| {
        t.token_type == TokenType::Identifier && t.value == method_name
    }).map(|offset| start + offset)
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

fn parse_expression_or_assignment(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let non_ws: Vec<&Token> = tokens.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

    // Response.Write expr
    if non_ws.len() >= 3
        && non_ws[0].value.eq_ignore_ascii_case("response")
        && non_ws[1].token_type == TokenType::Dot
        && non_ws[2].value.eq_ignore_ascii_case("write")
    {
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
        return Ok(Box::new(ResponseWrite::new(expr)));
    }

    // var = expr (bare assignment, no Set keyword)
    if non_ws.len() >= 2
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::Assign
    {
        let var_name = non_ws[0].value.clone();
        let assign_idx = tokens.iter().position(|t| t.token_type == TokenType::Assign).unwrap();
        let expr = parse_expression(&tokens[assign_idx + 1..])?;
        return Ok(Box::new(Assignment::new(var_name, expr)));
    }

    // arr(idx) = expr (array element assignment)
    if non_ws.len() >= 4
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::LeftParen
    {
        let var_name = non_ws[0].value.clone();
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
            return Err(VBSErrorType::SyntaxError.into_error(
                "Unmatched parentheses in array assignment".to_string()
            ));
        }
        let index_expr = parse_expression(&tokens[paren_start..i])?;
        i += 1;
        while i < tokens.len() && tokens[i].token_type == TokenType::WhiteSpace {
            i += 1;
        }
        if i >= tokens.len() || tokens[i].token_type != TokenType::Assign {
            return Err(VBSErrorType::SyntaxError.into_error(
                "Expected '=' after array index".to_string()
            ));
        }
        i += 1;
        let value_expr = parse_expression(&tokens[i..])?;
        return Ok(Box::new(ArrayAssignment::new(var_name, index_expr, value_expr)));
    }

    // obj.Method arg1, arg2, ... (method call)
    if non_ws.len() >= 3
        && non_ws[0].token_type == TokenType::Identifier
        && non_ws[1].token_type == TokenType::Dot
        && non_ws[2].token_type == TokenType::Identifier
    {
        let object_name = non_ws[0].value.clone();
        let method_name = non_ws[2].value.clone();

        let args = if let Some(mi) = find_method_token(tokens, &method_name) {
            let arg_tokens = &tokens[mi + 1..];
            parse_comma_args(arg_tokens)?
        } else {
            Vec::new()
        };

        return Ok(Box::new(MethodCall::new(object_name, method_name, args)));
    }

    Err(VBSErrorType::NotImplementedError
        .into_error(format!("Unrecognized command: {}", tokens_to_string(tokens))))
}

fn parse_line_into_syntax(tokens: &[Token]) -> Result<Box<dyn VBSyntax>, VBSError> {
    let first_token = tokens.iter()
        .find(|t| t.token_type != TokenType::WhiteSpace)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("Empty statement".to_string()))?;

    match first_token.token_type {
        TokenType::Dim => {
            let var_names = parse_dim_statement(tokens)?;
            Ok(Box::new(Dim::new(var_names)))
        }
        TokenType::Set => {
            parse_assignment_statement(tokens)
        }
        TokenType::ReDim => {
            parse_redim_statement(tokens)
        }
        _ => {
            parse_expression_or_assignment(tokens)
        }
    }
}

// ===== Token-to-Expr helpers =====

fn parse_tokens_to_expr(tokens: &[Token]) -> Result<Expr, VBSError> {
    if tokens.is_empty() {
        return Ok(Expr::Literal(VBValue::Empty));
    }
    parse_expression(tokens)
}

// ===== Block parsing =====

pub fn parse_blocks(lines: &[Vec<Token>]) -> Result<Vec<BlockStatement>, VBSError> {
    let mut pos = 0;
    parse_blocks_inner(lines, &mut pos)
}

fn parse_blocks_inner(lines: &[Vec<Token>], pos: &mut usize) -> Result<Vec<BlockStatement>, VBSError> {
    let mut blocks = Vec::new();

    while *pos < lines.len() {
        let line = &lines[*pos];
        let first = first_non_ws(line);

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
            Some(t)
                if t.token_type == TokenType::End
                    || t.token_type == TokenType::Next
                    || t.token_type == TokenType::WEnd
                    || t.token_type == TokenType::Loop
                    || t.token_type == TokenType::ElseIf
                    || t.token_type == TokenType::Else
                    || t.token_type == TokenType::Function
                    || t.token_type == TokenType::Sub =>
            {
                break;
            }
            _ => {
                // Skip comment lines
                if line.iter().any(|t| t.token_type == TokenType::Comment ||
                    (t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("rem")))
                {
                    *pos += 1;
                    continue;
                }
                blocks.push(match parse_line_into_syntax(line) {
                    Ok(syntax) => BlockStatement::Syntax(syntax),
                    Err(_) => BlockStatement::Unrecognized(tokens_to_string(line)),
                });
                *pos += 1;
            }
        }
    }

    Ok(blocks)
}

fn parse_expr_from_range(tokens: &[Token], start: usize, end: usize) -> Result<Expr, VBSError> {
    let filtered: Vec<Token> = tokens[start..end].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();
    parse_tokens_to_expr(&filtered)
}

fn parse_expr_from_slice(tokens: &[Token], start: usize) -> Result<Expr, VBSError> {
    let filtered: Vec<Token> = tokens[start..].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();
    parse_tokens_to_expr(&filtered)
}

fn parse_if_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    *pos += 1;

    let then_idx = find_keyword_or_type(line, "then", TokenType::Then)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("If without Then".to_string()))?;

    let condition = parse_expr_from_range(line, 1, then_idx)?;

    let after_then: Vec<&Token> = line[then_idx + 1..].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if !after_then.is_empty() {
        let inline_tokens: Vec<Token> = line[then_idx + 1..].to_vec();
        let line_text = tokens_to_string(&inline_tokens);
        let syntax = match parse_line_into_syntax(&inline_tokens) {
            Ok(s) => s,
            Err(_) => return Ok(BlockStatement::If {
                condition,
                then_body: vec![BlockStatement::Unrecognized(line_text)],
                else_if_blocks: Vec::new(),
                else_body: None,
            }),
        };
        let then_body = vec![BlockStatement::Syntax(syntax)];
        return Ok(BlockStatement::If {
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
                let second = next_line.iter()
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
                let syntax = parse_line_into_syntax(next_line).unwrap_or_else(|_| Box::new(create_error_syntax(line_text.clone())));
                match &section {
                    Section::Then => then_body.push(BlockStatement::Syntax(syntax)),
                    Section::ElseIf(idx) => else_if_blocks[*idx].body.push(BlockStatement::Syntax(syntax)),
                    Section::Else => { if let Some(ref mut eb) = else_body { eb.push(BlockStatement::Syntax(syntax)); } }
                }
                *pos += 1;
            }
            Some(t) if t.token_type == TokenType::ElseIf => {
                let elseif_line = &lines[*pos];
                *pos += 1;

                let then_idx = find_keyword_or_type(elseif_line, "then", TokenType::Then)
                    .ok_or_else(|| VBSErrorType::SyntaxError.into_error("ElseIf without Then".to_string()))?;

                let elseif_cond = parse_expr_from_range(elseif_line, 1, then_idx)?;

                let after_then: Vec<&Token> = elseif_line[then_idx + 1..].iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .collect();

                if !after_then.is_empty() {
                    let inline_tokens: Vec<Token> = elseif_line[then_idx + 1..].to_vec();
                    let line_text = tokens_to_string(&inline_tokens);
                    let syntax = match parse_line_into_syntax(&inline_tokens) {
                        Ok(s) => s,
                        Err(_) => Box::new(create_error_syntax(line_text)),
                    };
                    let inline_body = vec![BlockStatement::Syntax(syntax)];
                    else_if_blocks.push(ElseIfBlock { condition: elseif_cond, body: inline_body });
                    else_body.get_or_insert_with(Vec::new);
                    section = Section::Else;
                } else {
                    else_if_blocks.push(ElseIfBlock { condition: elseif_cond, body: Vec::new() });
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
                    Section::Else => { if let Some(ref mut eb) = else_body { eb.extend(sub_blocks); } }
                }
            }
        }
    }

    Ok(BlockStatement::If {
        condition,
        then_body,
        else_if_blocks,
        else_body,
    })
}

fn parse_for_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    *pos += 1;

    let for_line_no_ws: Vec<&Token> = line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

    if for_line_no_ws.len() < 5 {
        return Err(VBSErrorType::SyntaxError.into_error("Invalid For statement".to_string()));
    }

    let counter = for_line_no_ws[1].value.clone();

    if counter.eq_ignore_ascii_case("each") {
        return parse_for_each_block(line, pos, lines, &for_line_no_ws);
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

    let step: Option<Expr> = step_idx.map(|si| parse_expr_from_slice(line, si + 1)).transpose()?;

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
        counter,
        start,
        end,
        step,
        body,
    })
}

fn parse_for_each_block(
    line: &[Token],
    pos: &mut usize,
    lines: &[Vec<Token>],
    for_line_no_ws: &[&Token],
) -> Result<BlockStatement, VBSError> {
    if for_line_no_ws.len() < 5 {
        return Err(VBSErrorType::SyntaxError.into_error("Invalid For Each statement".to_string()));
    }

    let element = for_line_no_ws[2].value.clone();

    let in_idx = line.iter().position(|t| {
        t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("in")
    }).ok_or_else(|| VBSErrorType::SyntaxError.into_error("For Each without In".to_string()))?;

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

    Ok(BlockStatement::ForEach { element, group, body })
}

fn parse_while_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
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
            Some(t) if t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("wend") => {
                *pos += 1;
                break;
            }
            _ => {
                let sub_blocks = parse_blocks_inner(lines, pos)?;
                body.extend(sub_blocks);
            }
        }
    }

    Ok(BlockStatement::While { condition, body })
}

fn parse_do_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = &lines[*pos];
    *pos += 1;

    let do_line_no_ws: Vec<&Token> = line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

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
            let until_idx = line.iter().position(|t| {
                t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("until")
            }).unwrap();
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

                let loop_no_ws: Vec<&Token> = loop_line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();
                if loop_no_ws.len() > 1 {
                    let second = &loop_no_ws[1];
                    if second.value.eq_ignore_ascii_case("while") || second.token_type == TokenType::While {
                        is_post_until = false;
                        let while_idx = find_keyword_or_type(loop_line, "while", TokenType::While).unwrap();
                        post_condition = Some(parse_expr_from_slice(loop_line, while_idx + 1)?);
                    } else if second.value.eq_ignore_ascii_case("until") {
                        is_post_until = true;
                        let until_idx = loop_line.iter().position(|t| {
                            t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("until")
                        }).unwrap();
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
            body,
            condition: pre_condition,
            is_until,
            is_post_test: false,
        })
    } else {
        Ok(BlockStatement::Do {
            body,
            condition: post_condition,
            is_until: is_post_until,
            is_post_test: true,
        })
    }
}

// ===== Execution =====

fn evaluate_condition(expr: &Expr, context: &ExecutionContext) -> Result<bool, VBSError> {
    let val = evaluate(expr, context)?;
    if matches!(val, VBValue::Array(_) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    Ok(match val {
        VBValue::Boolean(b) => b,
        VBValue::Number(n) => n != 0.0,
        VBValue::String(s) => !s.is_empty(),
        VBValue::Null | VBValue::Empty => false,
        VBValue::Array(_) => unreachable!(),
        VBValue::Object(_) => unreachable!(),
    })
}

fn evaluate_numeric(expr: &Expr, context: &ExecutionContext) -> Result<f64, VBSError> {
    let val = evaluate(expr, context)?;
    if matches!(val, VBValue::Array(_) | VBValue::Object(_)) {
        return Err(VBSErrorType::ValueError.into_error("Type mismatch".to_string()));
    }
    Ok(match val {
        VBValue::Number(n) => n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty => 0.0,
        VBValue::Array(_) => unreachable!(),
        VBValue::Object(_) => unreachable!(),
    })
}

struct ErrorSyntax {
    message: String,
}

impl VBSyntax for ErrorSyntax {
    fn execute(&self, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::NotImplementedError
            .into_error(format!("Unrecognized command: {}", self.message)))
    }
}

fn create_error_syntax(message: String) -> ErrorSyntax {
    ErrorSyntax { message }
}

pub fn execute_blocks(
    blocks: &[BlockStatement],
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    for block in blocks {
        match block {
            BlockStatement::Syntax(syntax) => {
                syntax.execute(context)?;
            }
            BlockStatement::Unrecognized(line_text) => {
                return Err(VBSErrorType::NotImplementedError
                    .into_error(format!("Unrecognized command: {}", line_text)));
            }
            BlockStatement::If {
                condition,
                then_body,
                else_if_blocks,
                else_body,
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
            BlockStatement::For {
                counter,
                start,
                end,
                step,
                body,
            } => {
                let start_val = evaluate_numeric(start, context)?;
                let end_val = evaluate_numeric(end, context)?;
                let step_val = step.as_ref()
                    .map(|s| evaluate_numeric(s, context))
                    .unwrap_or(Ok(1.0))?;

                let mut i = start_val;
                if step_val > 0.0 {
                    while i <= end_val {
                        context.set_variable(counter, VBValue::Number(i));
                        execute_blocks(body, context)?;
                        i += step_val;
                    }
                } else if step_val < 0.0 {
                    while i >= end_val {
                        context.set_variable(counter, VBValue::Number(i));
                        execute_blocks(body, context)?;
                        i += step_val;
                    }
                }
                context.set_variable(counter, VBValue::Number(i));
            }
            BlockStatement::ForEach { element, group, body } => {
                let group_val = evaluate(group, context)?;
                match group_val {
                    VBValue::Array(ref items) => {
                        for item in items.iter() {
                            context.set_variable(element, item.clone());
                            execute_blocks(body, context)?;
                        }
                    }
                    _ => {
                        return Err(VBSErrorType::RuntimeError.into_error(
                            "Object doesn't support this property or method".to_string()
                        ));
                    }
                }
            }
            BlockStatement::While { condition, body } => {
                while evaluate_condition(condition, context)? {
                    execute_blocks(body, context)?;
                }
            }
            BlockStatement::Do {
                body,
                condition,
                is_until,
                is_post_test,
            } => {
                if *is_post_test {
                    loop {
                        execute_blocks(body, context)?;
                        if let Some(cond) = condition {
                            let result = evaluate_condition(cond, context)?;
                            if *is_until {
                                if result { break; }
                            } else {
                                if !result { break; }
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
                                if result { break; }
                            } else {
                                if !result { break; }
                            }
                        }
                        execute_blocks(body, context)?;
                    }
                }
            }
        }
    }
    Ok(())
}
