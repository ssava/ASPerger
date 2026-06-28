use super::block_parse::{first_non_ws, parse_blocks};
use super::block_types::{BlockStatement, CaseClause, UserDefinedFunction};
use crate::vbscript::execution_context::{ClassDefinition, ErrorMode, MethodDef, PropertyDef};
use crate::vbscript::expr::{evaluate, BinOp, Expr};
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::{ExecutionContext, Token, TokenType, VBValue};
use ahash::AHashMap;

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

pub(crate) fn execute_user_defined_function(
    func: &UserDefinedFunction,
    args: &[VBValue],
    context: &mut ExecutionContext,
) -> Result<VBValue, VBSError> {
    if let Some(ref debugger) = context.debugger {
        let vars: AHashMap<String, VBValue> = func
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| (p.to_lowercase(), args.get(i).cloned().unwrap_or(VBValue::Empty)))
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

    let _guard =
        crate::vbscript::execution_context::CodeStartLineGuard::new(&mut context.code_start_line);

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
        Err(e) => return Err(e),
    }

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
                    .map(|t| {
                        t.token_type == TokenType::Get || t.value.eq_ignore_ascii_case("get")
                    })
                    .unwrap_or(false);
                let is_let = get_let_set
                    .map(|t| {
                        t.token_type == TokenType::Let || t.value.eq_ignore_ascii_case("let")
                    })
                    .unwrap_or(false);
                if is_get || is_let {
                    let name_tok = match name_tok {
                        Some(t) if t.token_type == TokenType::Identifier => t,
                        _ => {
                            i += 1;
                            continue;
                        }
                    };
                    let prop_name = name_tok.value.to_lowercase();
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
                                    param = Some(param_tok.value.to_lowercase());
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
                        .entry(prop_name.clone())
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
        let method_name = no_ws[name_idx].value.to_lowercase();

        let mut params = Vec::new();
        if no_ws.len() > name_idx + 1 && no_ws[name_idx + 1].token_type == TokenType::LeftParen {
            let mut p = name_idx + 2;
            while p < no_ws.len() && no_ws[p].token_type != TokenType::RightParen {
                if no_ws[p].token_type == TokenType::Identifier {
                    params.push(no_ws[p].value.to_lowercase());
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

fn execute_for_block(
    body: &[BlockStatement],
    counter: &str,
    start: &Expr,
    end: &Expr,
    step: &Option<Expr>,
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
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
    Ok(())
}

fn execute_foreach_block(
    element: &str,
    group: &Expr,
    body: &[BlockStatement],
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
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
    Ok(())
}

fn execute_do_block(
    body: &[BlockStatement],
    condition: &Option<Expr>,
    is_until: bool,
    is_post_test: bool,
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    if is_post_test {
        loop {
            match execute_blocks(body, context) {
                Ok(()) => {}
                Err(e) if e.is_exit_do() => break,
                Err(e) => return Err(e),
            }
            if let Some(cond) = condition {
                let result = evaluate_condition(cond, context)?;
                if is_until {
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
                if is_until {
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
    Ok(())
}

fn execute_select_case_block(
    expression: &Expr,
    cases: &[CaseClause],
    else_body: &Option<Vec<BlockStatement>>,
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    let expr_val = evaluate(expression, context)?;
    context.select_value = Some(expr_val.clone());
    let mut matched = false;

    for case in cases {
        for val_expr in &case.values {
            let case_val = evaluate(val_expr, context)?;
            let is_match = match val_expr {
                Expr::CaseComparison { .. } => matches!(case_val, VBValue::Boolean(true)),
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
    Ok(())
}

pub fn execute_blocks(
    blocks: &[BlockStatement],
    context: &mut ExecutionContext,
) -> Result<(), VBSError> {
    tracing::trace!(block_count = blocks.len(), "Executing VB blocks");
    for block in blocks {
        tracing::trace!(line = block.line(), "Executing block");
        if context.response.ended {
            break;
        }

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
            BlockStatement::Syntax(syntax, _line) => {
                let result = syntax.execute(context);
                if let Err(e) = result {
                    if *context.get_error_mode() == ErrorMode::ResumeNext {
                        context.set_err(e);
                    } else {
                        return Err(e);
                    }
                }
            }
            BlockStatement::Unrecognized(err, _line_text, _line) => {
                return Err(err.clone());
            }
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
            BlockStatement::For { counter, start, end, step, body, .. } => {
                execute_for_block(body, counter, start, end, step, context)?;
            }
            BlockStatement::ForEach { element, group, body, .. } => {
                execute_foreach_block(element, group, body, context)?;
            }
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
            BlockStatement::Do { body, condition, is_until, is_post_test, .. } => {
                execute_do_block(body, condition, *is_until, *is_post_test, context)?;
            }
            BlockStatement::SelectCase { expression, cases, else_body, .. } => {
                execute_select_case_block(expression, cases, else_body, context)?;
            }
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
            BlockStatement::With { object, body, .. } => {
                let obj_val = evaluate(object, context)?;
                let prev_with = context.with_object.replace(obj_val);
                let result = execute_blocks(body, context);
                context.with_object = prev_with;
                result?
            }
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
