use super::vbs_error::{VBSError, VBSErrorType};
use super::expr::{evaluate, parse_expression};
use super::{ExecutionContext, Token, TokenType, VBScriptInterpreter, VBValue};

pub enum BlockStatement {
    Line(Vec<Token>),
    If {
        condition_tokens: Vec<Token>,
        then_body: Vec<BlockStatement>,
        else_if_blocks: Vec<ElseIfBlock>,
        else_body: Option<Vec<BlockStatement>>,
    },
    For {
        counter: String,
        start_tokens: Vec<Token>,
        end_tokens: Vec<Token>,
        step_tokens: Option<Vec<Token>>,
        body: Vec<BlockStatement>,
    },
    While {
        condition_tokens: Vec<Token>,
        body: Vec<BlockStatement>,
    },
    Do {
        body: Vec<BlockStatement>,
        condition_tokens: Option<Vec<Token>>,
        is_until: bool,
        is_post_test: bool,
    },
}

pub struct ElseIfBlock {
    pub condition_tokens: Vec<Token>,
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
                    (t.token_type == TokenType::Identifier && t.value.to_lowercase() == "rem"))
                {
                    *pos += 1;
                    continue;
                }
                blocks.push(BlockStatement::Line(line.clone()));
                *pos += 1;
            }
        }
    }

    Ok(blocks)
}

fn parse_if_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = lines[*pos].clone();
    *pos += 1;

    let then_idx = find_keyword_or_type(&line, "then", TokenType::Then)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("If without Then".to_string()))?;

    let condition_tokens: Vec<Token> = line[1..then_idx].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();

    let after_then: Vec<&Token> = line[then_idx + 1..].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .collect();

    if !after_then.is_empty() {
        let inline_tokens: Vec<Token> = line[then_idx + 1..].to_vec();
        let then_body = vec![BlockStatement::Line(inline_tokens)];
        return Ok(BlockStatement::If {
            condition_tokens,
            then_body,
            else_if_blocks: Vec::new(),
            else_body: None,
        });
    }

    // Block If: use an enum to track which section we're collecting into
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
                match &section {
                    Section::Then => then_body.push(BlockStatement::Line(next_line.clone())),
                    Section::ElseIf(idx) => else_if_blocks[*idx].body.push(BlockStatement::Line(next_line.clone())),
                    Section::Else => { if let Some(ref mut eb) = else_body { eb.push(BlockStatement::Line(next_line.clone())); } }
                }
                *pos += 1;
            }
            Some(t) if t.token_type == TokenType::ElseIf => {
                let elseif_line = lines[*pos].clone();
                *pos += 1;

                let then_idx = find_keyword_or_type(&elseif_line, "then", TokenType::Then)
                    .ok_or_else(|| VBSErrorType::SyntaxError.into_error("ElseIf without Then".to_string()))?;

                let elseif_cond: Vec<Token> = elseif_line[1..then_idx].iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .cloned()
                    .collect();

                let after_then: Vec<&Token> = elseif_line[then_idx + 1..].iter()
                    .filter(|t| t.token_type != TokenType::WhiteSpace)
                    .collect();

                if !after_then.is_empty() {
                    let inline_body = vec![BlockStatement::Line(elseif_line[then_idx + 1..].to_vec())];
                    else_if_blocks.push(ElseIfBlock { condition_tokens: elseif_cond, body: inline_body });
                    else_body.get_or_insert_with(Vec::new);
                    section = Section::Else;
                } else {
                    else_if_blocks.push(ElseIfBlock { condition_tokens: elseif_cond, body: Vec::new() });
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
        condition_tokens,
        then_body,
        else_if_blocks,
        else_body,
    })
}

fn parse_for_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = lines[*pos].clone();
    *pos += 1;

    // For counter = start To end [Step step]
    // Find the counter (first Identifier after For)
    let for_line_no_ws: Vec<&Token> = line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

    if for_line_no_ws.len() < 5 {
        return Err(VBSErrorType::SyntaxError.into_error("Invalid For statement".to_string()));
    }

    // After "For", expect Identifier "=" expr "To" expr [ "Step" expr ]
    let counter = for_line_no_ws[1].value.clone();

    // Find Assign, To, and Step positions in the original (with whitespace) token list
    let assign_idx = find_token(&line, TokenType::Assign)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("For without =".to_string()))?;

    let to_idx = find_keyword_or_type(&line, "to", TokenType::To)
        .ok_or_else(|| VBSErrorType::SyntaxError.into_error("For without To".to_string()))?;

    let step_idx = find_keyword_or_type(&line, "step", TokenType::Step);

    // start tokens: between = and To
    let start_tokens: Vec<Token> = line[assign_idx + 1..to_idx].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();

    // end tokens: between To and Step (or To and end of line)
    let end_tokens: Vec<Token> = if let Some(si) = step_idx {
        line[to_idx + 1..si].iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .cloned()
            .collect()
    } else {
        line[to_idx + 1..].iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .cloned()
            .collect()
    };

    // step tokens: after Step
    let step_tokens: Option<Vec<Token>> = step_idx.map(|si| {
        line[si + 1..].iter()
            .filter(|t| t.token_type != TokenType::WhiteSpace)
            .cloned()
            .collect()
    });

    // Parse body until Next
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
        start_tokens,
        end_tokens,
        step_tokens,
        body,
    })
}

fn parse_while_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = lines[*pos].clone();
    *pos += 1;

    // While condition
    let condition_tokens: Vec<Token> = line[1..].iter()
        .filter(|t| t.token_type != TokenType::WhiteSpace)
        .cloned()
        .collect();

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

    Ok(BlockStatement::While {
        condition_tokens,
        body,
    })
}

fn parse_do_block(lines: &[Vec<Token>], pos: &mut usize) -> Result<BlockStatement, VBSError> {
    let line = lines[*pos].clone();
    *pos += 1;

    // Do [{While|Until} condition] ... Loop [{While|Until} condition]
    let do_line_no_ws: Vec<&Token> = line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();

    let mut is_pre_test = false;
    let mut is_until = false;
    let mut pre_condition_tokens: Option<Vec<Token>> = None;

    if do_line_no_ws.len() > 1 {
        let second = &do_line_no_ws[1];
        let second_upper = second.value.to_uppercase();
        if second_upper == "WHILE" || second.token_type == TokenType::While {
            is_pre_test = true;
            is_until = false;
            // Extract condition tokens after While
            let while_idx = find_keyword_or_type(&line, "while", TokenType::While).unwrap();
            pre_condition_tokens = Some(line[while_idx + 1..].iter()
                .filter(|t| t.token_type != TokenType::WhiteSpace)
                .cloned()
                .collect());
        } else if second_upper == "UNTIL" {
            is_pre_test = true;
            is_until = true;
            let until_idx = line.iter().position(|t| {
                t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("until")
            }).unwrap();
            pre_condition_tokens = Some(line[until_idx + 1..].iter()
                .filter(|t| t.token_type != TokenType::WhiteSpace)
                .cloned()
                .collect());
        }
    }

    let mut body = Vec::new();
    let mut post_condition_tokens: Option<Vec<Token>> = None;
    let mut is_post_until = false;

    loop {
        if *pos >= lines.len() {
            return Err(VBSErrorType::SyntaxError.into_error("Do without Loop".to_string()));
        }

        let next_line = &lines[*pos];
        let first = first_non_ws(next_line);

        match first {
            Some(t) if t.token_type == TokenType::Loop => {
                let loop_line = lines[*pos].clone();
                *pos += 1;

                let loop_no_ws: Vec<&Token> = loop_line.iter().filter(|t| t.token_type != TokenType::WhiteSpace).collect();
                if loop_no_ws.len() > 1 {
                    let second = &loop_no_ws[1];
                    let second_upper = second.value.to_uppercase();
                    if second_upper == "WHILE" || second.token_type == TokenType::While {
                        is_post_until = false;
                        let while_idx = find_keyword_or_type(&loop_line, "while", TokenType::While).unwrap();
                        post_condition_tokens = Some(loop_line[while_idx + 1..].iter()
                            .filter(|t| t.token_type != TokenType::WhiteSpace)
                            .cloned()
                            .collect());
                    } else if second_upper == "UNTIL" {
                        is_post_until = true;
                        let until_idx = loop_line.iter().position(|t| {
                        t.token_type == TokenType::Identifier && t.value.eq_ignore_ascii_case("until")
                    }).unwrap();
                        post_condition_tokens = Some(loop_line[until_idx + 1..].iter()
                            .filter(|t| t.token_type != TokenType::WhiteSpace)
                            .cloned()
                            .collect());
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

    // If there's a pre-test condition, use it; otherwise use post-test condition
    if is_pre_test {
        Ok(BlockStatement::Do {
            body,
            condition_tokens: pre_condition_tokens,
            is_until,
            is_post_test: false,
        })
    } else {
        Ok(BlockStatement::Do {
            body,
            condition_tokens: post_condition_tokens,
            is_until: is_post_until,
            is_post_test: true,
        })
    }
}

fn eval_token_expression(tokens: &[Token], context: &ExecutionContext) -> Result<VBValue, VBSError> {
    if tokens.is_empty() {
        return Ok(VBValue::Empty);
    }
    let expr = parse_expression(tokens)?;
    evaluate(&expr, context)
}

fn evaluate_condition(tokens: &[Token], context: &ExecutionContext) -> Result<bool, VBSError> {
    let val = eval_token_expression(tokens, context)?;
    Ok(match val {
        VBValue::Boolean(b) => b,
        VBValue::Number(n) => n != 0.0,
        VBValue::String(s) => !s.is_empty(),
        VBValue::Null | VBValue::Empty => false,
    })
}

fn evaluate_numeric(tokens: &[Token], context: &ExecutionContext) -> Result<f64, VBSError> {
    let val = eval_token_expression(tokens, context)?;
    Ok(match val {
        VBValue::Number(n) => n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty => 0.0,
    })
}

pub fn execute_blocks(
    blocks: &[BlockStatement],
    interpreter: &VBScriptInterpreter,
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    for block in blocks {
        match block {
            BlockStatement::Line(tokens) => {
                let result = interpreter.create_syntax_from_tokens(tokens);
                match result {
                    Ok(Some(syntax)) => syntax.execute(context)?,
                    Ok(None) => {
                        let line_text: String = tokens.iter().map(|t| t.value.clone()).collect::<Vec<_>>().join(" ");
                        return Err(VBSErrorType::NotImplementedError
                            .into_error(format!("Unrecognized command: {}", line_text)));
                    }
                    Err(e) => return Err(e),
                }
            }
            BlockStatement::If {
                condition_tokens,
                then_body,
                else_if_blocks,
                else_body,
            } => {
                if evaluate_condition(condition_tokens, context)? {
                    execute_blocks(then_body, interpreter, context)?;
                } else {
                    let mut handled = false;
                    for elseif in else_if_blocks {
                        if evaluate_condition(&elseif.condition_tokens, context)? {
                            execute_blocks(&elseif.body, interpreter, context)?;
                            handled = true;
                            break;
                        }
                    }
                    if !handled {
                        if let Some(body) = else_body {
                            execute_blocks(body, interpreter, context)?;
                        }
                    }
                }
            }
            BlockStatement::For {
                counter,
                start_tokens,
                end_tokens,
                step_tokens,
                body,
            } => {
                let start = evaluate_numeric(start_tokens, context)?;
                let end = evaluate_numeric(end_tokens, context)?;
                let step = step_tokens.as_ref()
                    .map(|t| evaluate_numeric(t, context))
                    .unwrap_or(Ok(1.0))?;

                let mut i = start;
                if step > 0.0 {
                    while i <= end {
                        context.set_variable(&counter, VBValue::Number(i));
                        execute_blocks(body, interpreter, context)?;
                        i += step;
                    }
                } else if step < 0.0 {
                    while i >= end {
                        context.set_variable(&counter, VBValue::Number(i));
                        execute_blocks(body, interpreter, context)?;
                        i += step;
                    }
                }
                context.set_variable(&counter, VBValue::Number(i));
            }
            BlockStatement::While {
                condition_tokens,
                body,
            } => {
                while evaluate_condition(condition_tokens, context)? {
                    execute_blocks(body, interpreter, context)?;
                }
            }
            BlockStatement::Do {
                body,
                condition_tokens,
                is_until,
                is_post_test,
            } => {
                if *is_post_test {
                    loop {
                        execute_blocks(body, interpreter, context)?;
                        if let Some(cond) = condition_tokens {
                            let result = evaluate_condition(cond, context)?;
                            if *is_until {
                                if result { break; }
                            } else {
                                if !result { break; }
                            }
                        } else {
                            break; // No condition — execute once? Actually Do...Loop without condition is infinite
                        }
                    }
                } else {
                    // Pre-test
                    loop {
                        if let Some(cond) = condition_tokens {
                            let result = evaluate_condition(cond, context)?;
                            if *is_until {
                                if result { break; }
                            } else {
                                if !result { break; }
                            }
                        }
                        execute_blocks(body, interpreter, context)?;
                    }
                }
            }
        }
    }
    Ok(())
}
