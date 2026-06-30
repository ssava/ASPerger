use crate::vbscript::block::first_non_ws;
use crate::vbscript::block::BlockStatement;
use crate::vbscript::block::UserDefinedFunction;
use crate::vbscript::execution_context::{ClassDefinition, MethodDef, PropertyDef};
use crate::vbscript::expr::Expr;
use crate::vbscript::instruction::Instruction;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::{ExecutionContext, Token, TokenType, VBValue};
use ahash::AHashMap;

#[derive(Clone)]
pub struct CompiledCode {
    pub instructions: Vec<Instruction>,
    pub constants: Vec<VBValue>,
    pub local_count: usize,
    pub local_names: Vec<String>,
    pub function_defs: Vec<UserDefinedFunction>,
    pub compiled_functions: Vec<(String, CompiledCode)>,
}

pub struct Compiler<'a> {
    code: Vec<Instruction>,
    constants: Vec<VBValue>,
    locals: ahash::AHashMap<String, usize>,
    local_count: usize,
    context: &'a mut ExecutionContext,
    function_defs: Vec<UserDefinedFunction>,
    loop_stack: Vec<LoopInfo>,
    compiled_functions: ahash::AHashMap<String, CompiledCode>,
}

#[allow(dead_code)]
struct LoopInfo {
    break_target: usize,
    continue_target: usize,
    exits: Vec<Patch>,
}

struct Patch(usize);

impl<'a> Compiler<'a> {
    pub fn new(context: &'a mut ExecutionContext) -> Self {
        Compiler {
            code: Vec::new(),
            constants: Vec::new(),
            locals: ahash::AHashMap::new(),
            local_count: 0,
            context,
            loop_stack: Vec::new(),
            compiled_functions: ahash::AHashMap::new(),
            function_defs: Vec::new(),
        }
    }

    pub fn compile(&mut self, blocks: &[BlockStatement]) -> Result<CompiledCode, VBSError> {
        self.compile_blocks(blocks)?;
        let mut local_names = vec![String::new(); self.local_count];
        for (name, slot) in &self.locals {
            if *slot < local_names.len() {
                local_names[*slot] = name.clone();
            }
        }
        Ok(CompiledCode {
            instructions: std::mem::take(&mut self.code),
            constants: std::mem::take(&mut self.constants),
            local_count: self.local_count,
            local_names,
            function_defs: std::mem::take(&mut self.function_defs),
            compiled_functions: self.compiled_functions.drain().collect(),
        })
    }

    pub(crate) fn emit(&mut self, inst: Instruction) {
        self.code.push(inst);
    }

    fn current_offset(&self) -> usize {
        self.code.len()
    }

    fn emit_jump(&mut self) -> Patch {
        let pos = self.code.len();
        self.code.push(Instruction::Jump(0));
        Patch(pos)
    }

    fn emit_jump_if_false(&mut self) -> Patch {
        let pos = self.code.len();
        self.code.push(Instruction::JumpIfFalse(0));
        Patch(pos)
    }

    fn emit_jump_if_true(&mut self) -> Patch {
        let pos = self.code.len();
        self.code.push(Instruction::JumpIfTrue(0));
        Patch(pos)
    }

    fn patch_jump(&mut self, patch: Patch, target: usize) {
        let offset = (target as isize - patch.0 as isize - 1) as i32;
        match &mut self.code[patch.0] {
            Instruction::Jump(o) | Instruction::JumpIfFalse(o) | Instruction::JumpIfTrue(o) => {
                *o = offset;
            }
            Instruction::ForPrep(_, o) | Instruction::ForEachPrep(_, o) => {
                *o = offset;
            }
            _ => unreachable!(),
        }
    }

    pub(crate)     fn add_constant(&mut self, val: VBValue) -> u32 {
        let idx = self.constants.len();
        self.constants.push(val);
        idx as u32
    }

    pub(crate) fn allocate_local(&mut self, name: &str) -> usize {
        let key = name.to_lowercase();
        if let Some(&slot) = self.locals.get(&key) {
            return slot;
        }
        let slot = self.local_count;
        self.locals.insert(key, slot);
        self.local_count += 1;
        slot
    }

    pub(crate) fn local_slot(&self, name: &str) -> Option<usize> {
        self.locals.get(&name.to_lowercase()).copied()
    }

    fn compile_blocks(&mut self, blocks: &[BlockStatement]) -> Result<(), VBSError> {
        for block in blocks {
            self.compile_block(block)?;
        }
        Ok(())
    }

    fn compile_block(&mut self, block: &BlockStatement) -> Result<(), VBSError> {
        match block {
            BlockStatement::Syntax(syntax, _line) => {
                syntax.compile(self)?;
            }
            BlockStatement::If {
                condition,
                then_body,
                else_if_blocks,
                else_body,
                ..
            } => {
                self.compile_expr(condition);
                let else_jump = self.emit_jump_if_false();

                for b in then_body {
                    self.compile_block(b)?;
                }

                let mut end_patches = Vec::new();
                if !else_if_blocks.is_empty() || else_body.is_some() {
                    let end_jump = self.emit_jump();
                    end_patches.push(end_jump);
                }

                self.patch_jump(else_jump, self.current_offset());

                for elseif in else_if_blocks {
                    self.compile_expr(&elseif.condition);
                    let next_elseif = self.emit_jump_if_false();

                    for b in &elseif.body {
                        self.compile_block(b)?;
                    }

                    let end_jump = self.emit_jump();
                    end_patches.push(end_jump);
                    self.patch_jump(next_elseif, self.current_offset());
                }

                if let Some(else_body) = else_body {
                    for b in else_body {
                        self.compile_block(b)?;
                    }
                }

                for p in end_patches {
                    self.patch_jump(p, self.current_offset());
                }
            }
            BlockStatement::For {
                counter,
                start,
                end,
                step,
                body,
                ..
            } => {
                let slot = self.allocate_local(counter);
                self.compile_expr(start);
                self.emit(Instruction::StoreLocal(slot));

                self.compile_expr(end);
                self.compile_expr(step.as_ref().unwrap_or(&Expr::Literal(VBValue::Number(1.0))));

                let prep_offset = self.current_offset();
                let break_target = prep_offset;
                self.emit(Instruction::ForPrep(slot, 0));
                let exit_patch = Patch(self.code.len() - 1);

                let continue_target = self.current_offset();
                self.loop_stack.push(LoopInfo {
                    break_target,
                    continue_target,
                    exits: Vec::new(),
                });

                for b in body {
                    self.compile_block(b)?;
                }

                let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);

                self.emit(Instruction::ForStep(slot, (continue_target as isize - self.current_offset() as isize - 1) as i32));

                let after_loop = self.current_offset();
                self.patch_jump(exit_patch, after_loop);
                for exit in exits {
                    self.patch_jump(exit, after_loop);
                }
            }
            BlockStatement::ForEach {
                element,
                group,
                body,
                ..
            } => {
                let slot = self.allocate_local(element);
                self.compile_expr(group);

                let prep_offset = self.current_offset();
                let break_target = prep_offset;
                self.emit(Instruction::ForEachPrep(slot, 0));
                let exit_patch = Patch(self.code.len() - 1);

                let continue_target = self.current_offset();
                self.loop_stack.push(LoopInfo {
                    break_target,
                    continue_target,
                    exits: Vec::new(),
                });

                for b in body {
                    self.compile_block(b)?;
                }

                let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);

                self.emit(
                    Instruction::ForEachStep(slot, (continue_target as isize - self.current_offset() as isize - 1) as i32),
                );

                let after_loop = self.current_offset();
                self.patch_jump(exit_patch, after_loop);
                for exit in exits {
                    self.patch_jump(exit, after_loop);
                }
            }
            BlockStatement::While { condition, body, .. } => {
                let loop_start = self.current_offset();
                let break_target = loop_start;
                self.compile_expr(condition);
                let exit_patch = self.emit_jump_if_false();

                let continue_target = self.current_offset();
                self.loop_stack.push(LoopInfo {
                    break_target,
                    continue_target,
                    exits: Vec::new(),
                });

                for b in body {
                    self.compile_block(b)?;
                }

                let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);

                self.emit(Instruction::Jump((loop_start as isize - self.current_offset() as isize - 1) as i32));
                let after_loop = self.current_offset();
                self.patch_jump(exit_patch, after_loop);
                for exit in exits {
                    self.patch_jump(exit, after_loop);
                }
            }
            BlockStatement::Do {
                body,
                condition,
                is_until,
                is_post_test,
                ..
            } => {
                if *is_post_test {
                    let loop_start = self.current_offset();
                    let break_target = loop_start;
                    self.loop_stack.push(LoopInfo {
                        break_target,
                        continue_target: loop_start,
                        exits: Vec::new(),
                    });

                    for b in body {
                        self.compile_block(b)?;
                    }

                    let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);

                    if let Some(cond) = condition {
                        self.compile_expr(cond);
                        if *is_until {
                            self.emit(Instruction::JumpIfFalse(
                                (loop_start as isize - self.current_offset() as isize - 1) as i32,
                            ));
                        } else {
                            self.emit(Instruction::JumpIfTrue(
                                (loop_start as isize - self.current_offset() as isize - 1) as i32,
                            ));
                        }
                    } else {
                        self.emit(Instruction::Jump(
                            (loop_start as isize - self.current_offset() as isize - 1) as i32,
                        ));
                    }
                    let after_loop = self.current_offset();
                    for exit in exits {
                        self.patch_jump(exit, after_loop);
                    }
                } else {
                    let loop_start = self.current_offset();
                    let break_target = loop_start;

                    if let Some(cond) = condition {
                        self.compile_expr(cond);
                        if *is_until {
                            let exit_patch = self.emit_jump_if_true();
                            let continue_target = self.current_offset();
                            self.loop_stack.push(LoopInfo {
                                break_target,
                                continue_target,
                                exits: Vec::new(),
                            });

                            for b in body {
                                self.compile_block(b)?;
                            }

                            let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);
                            self.emit(Instruction::Jump(
                                (loop_start as isize - self.current_offset() as isize - 1) as i32,
                            ));
                            let after_loop = self.current_offset();
                            self.patch_jump(exit_patch, after_loop);
                            for exit in exits {
                                self.patch_jump(exit, after_loop);
                            }
                        } else {
                            let exit_patch = self.emit_jump_if_false();
                            let continue_target = self.current_offset();
                            self.loop_stack.push(LoopInfo {
                                break_target,
                                continue_target,
                                exits: Vec::new(),
                            });

                            for b in body {
                                self.compile_block(b)?;
                            }

                            let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);
                            self.emit(Instruction::Jump(
                                (loop_start as isize - self.current_offset() as isize - 1) as i32,
                            ));
                            let after_loop = self.current_offset();
                            self.patch_jump(exit_patch, after_loop);
                            for exit in exits {
                                self.patch_jump(exit, after_loop);
                            }
                        }
                    } else {
                        let continue_target = self.current_offset();
                        self.loop_stack.push(LoopInfo {
                            break_target,
                            continue_target,
                            exits: Vec::new(),
                        });

                        for b in body {
                            self.compile_block(b)?;
                        }

                        let exits = self.loop_stack.pop().map_or(Vec::new(), |info| info.exits);
                        self.emit(Instruction::Jump(
                            (loop_start as isize - self.current_offset() as isize - 1) as i32,
                        ));
                        let after_loop = self.current_offset();
                        for exit in exits {
                            self.patch_jump(exit, after_loop);
                        }
                    }
                }
            }
            BlockStatement::SelectCase {
                expression,
                cases,
                else_body,
                ..
            } => {
                self.compile_expr(expression);
                self.emit(Instruction::SelectStore);

                let mut end_patches = Vec::new();

                for case in cases {
                    for val_expr in &case.values {
                        self.compile_expr(val_expr);
                        // CaseComparison and Range already push a boolean, skip SelectCompare
                        let needs_compare = !matches!(val_expr, Expr::CaseComparison { .. } | Expr::Range { .. });
                        if needs_compare {
                            self.emit(Instruction::SelectCompare);
                        }
                        let next_patch = self.emit_jump_if_false();

                        for b in &case.body {
                            self.compile_block(b)?;
                        }
                        let end_jump = self.emit_jump();
                        end_patches.push(end_jump);
                        self.patch_jump(next_patch, self.current_offset());
                    }
                }

                if let Some(else_body) = else_body {
                    for b in else_body {
                        self.compile_block(b)?;
                    }
                }

                self.emit(Instruction::SelectClear);
                for p in end_patches {
                    self.patch_jump(p, self.current_offset());
                }
            }
            BlockStatement::ClassDef {
                name,
                body_lines,
                ..
            } => {
                if let Ok(properties) = extract_properties_from_class_body(body_lines) {
                    let methods = extract_methods_from_class_body(body_lines);
                    let class_def = ClassDefinition {
                        name: name.clone(),
                        properties,
                        methods,
                    };
                    self.context.define_class(class_def);
                    if let Some(class_def) = self.context.get_class(name) {
                        let prop_bodies: Vec<_> = class_def.properties.values().map(|pd| {
                            (pd.name.clone(), pd.get_body.clone(), pd.let_body.clone(), pd.let_param.clone())
                        }).collect();
                        let method_bodies: Vec<_> = class_def.methods.values().map(|md| {
                            (md.name.clone(), md.params.clone(), md.body_lines.clone())
                        }).collect();
                        for (prop_name, get_body, let_body, let_param) in &prop_bodies {
                            if let Some(body) = get_body {
                                let lines = self.token_lines_to_blocks(body);
                                let func_name = format!("__cls_{}_get_{}", name, prop_name);
                                let compiled = self.compile_function(&lines, &[])?;
                                self.compiled_functions
                                    .insert(func_name.to_lowercase(), compiled);
                            }
                            if let Some(body) = let_body {
                                let lines = self.token_lines_to_blocks(body);
                                let func_name = format!("__cls_{}_let_{}", name, prop_name);
                                let param = let_param.as_deref().unwrap_or("__value__");
                                let compiled = self.compile_function(&lines, &[param])?;
                                self.compiled_functions
                                    .insert(func_name.to_lowercase(), compiled);
                            }
                        }
                        for (method_name, params, body_lines) in &method_bodies {
                            let lines = self.token_lines_to_blocks(body_lines);
                            let func_name = format!("__cls_{}_{}", name, method_name);
                            let params: Vec<&str> = params.iter().map(|s| s.as_str()).collect();
                            let compiled = self.compile_function(&lines, &params)?;
                            self.compiled_functions
                                .insert(func_name.to_lowercase(), compiled);
                        }
                    }
                }
            }
            BlockStatement::With { object, body, .. } => {
                self.compile_expr(object);
                self.emit(Instruction::WithStart);
                for b in body {
                    self.compile_block(b)?;
                }
                self.emit(Instruction::WithEnd);
            }
            BlockStatement::FunctionDef {
                name,
                params,
                body_lines,
                ..
            } => {
                self.function_defs.push(UserDefinedFunction {
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                    is_function: true,
                });
                let lines = self.token_lines_to_blocks(body_lines);
                let param_strs: Vec<&str> = params.iter().map(|s| s.as_str()).collect();
                let compiled = self.compile_function(&lines, &param_strs)?;
                let key = name.to_lowercase();
                self.compiled_functions.insert(key, compiled);
            }
            BlockStatement::SubDef {
                name,
                params,
                body_lines,
                ..
            } => {
                self.function_defs.push(UserDefinedFunction {
                    name: name.clone(),
                    params: params.clone(),
                    body_lines: body_lines.clone(),
                    is_function: false,
                });
                let lines = self.token_lines_to_blocks(body_lines);
                let param_strs: Vec<&str> = params.iter().map(|s| s.as_str()).collect();
                let compiled = self.compile_function(&lines, &param_strs)?;
                let key = name.to_lowercase();
                self.compiled_functions.insert(key, compiled);
            }
            BlockStatement::ExitFor(_) | BlockStatement::ExitDo(_) => {
                self.emit_exit();
            }
            BlockStatement::ExitFunction(_) => {
                self.emit(Instruction::ExitFunction);
            }
            BlockStatement::ExitSub(_) => {
                self.emit(Instruction::ExitSub);
            }
            BlockStatement::Unrecognized(e, _msg, _line) => {
                return Err(e.clone());
            }
        }
        Ok(())
    }

    fn compile_function(
        &mut self,
        blocks: &[BlockStatement],
        params: &[&str],
    ) -> Result<CompiledCode, VBSError> {
        let saved_locals = self.locals.clone();
        let saved_local_count = self.local_count;
        let saved_code = std::mem::take(&mut self.code);
        let saved_constants = std::mem::take(&mut self.constants);
        let saved_loop_stack = std::mem::take(&mut self.loop_stack);
        let saved_compiled_functions = std::mem::take(&mut self.compiled_functions);

        self.locals.clear();
        self.local_count = 0;
        for p in params {
            self.allocate_local(p);
        }

        self.compile_blocks(blocks)?;

        let mut local_names = vec![String::new(); self.local_count];
        for (name, slot) in &self.locals {
            if *slot < local_names.len() {
                local_names[*slot] = name.clone();
            }
        }
        let compiled = CompiledCode {
            instructions: std::mem::take(&mut self.code),
            constants: std::mem::take(&mut self.constants),
            local_count: self.local_count,
            local_names,
            function_defs: Vec::new(),
            compiled_functions: Vec::new(),
        };

        self.code = saved_code;
        self.constants = saved_constants;
        self.locals = saved_locals;
        self.local_count = saved_local_count;
        self.loop_stack = saved_loop_stack;
        self.compiled_functions = saved_compiled_functions;

        Ok(compiled)
    }

    fn token_lines_to_blocks(&self, body_lines: &[Vec<crate::vbscript::Token>]) -> Vec<BlockStatement> {
        crate::vbscript::block::parse_blocks(body_lines).unwrap_or_default()
    }

    pub(crate) fn emit_exit(&mut self) {
        let pos = self.code.len();
        self.emit(Instruction::Jump(0));
        if let Some(loop_info) = self.loop_stack.last_mut() {
            loop_info.exits.push(Patch(pos));
        }
    }

    pub(crate) fn compile_expr(&mut self, expr: &Expr) {
        match expr {
            Expr::Literal(val) => {
                let idx = self.add_constant(val.clone());
                self.emit(Instruction::LoadConst(idx));
            }
            Expr::Variable(name) => {
                let name_lower = name.to_lowercase();
                if self.locals.contains_key(&name_lower) {
                    let slot = self.locals[&name_lower];
                    self.emit(Instruction::LoadLocal(slot));
                } else {
                    let idx = self.add_constant(VBValue::String(name_lower.into()));
                    self.emit(Instruction::LoadGlobal(idx));
                }
            }
            Expr::BinaryOp { left, op, right } => {
                self.compile_expr(left);
                self.compile_expr(right);
                let inst = match op {
                    crate::vbscript::expr::BinOp::Add => Instruction::Add,
                    crate::vbscript::expr::BinOp::Sub => Instruction::Sub,
                    crate::vbscript::expr::BinOp::Mul => Instruction::Mul,
                    crate::vbscript::expr::BinOp::Div => Instruction::Div,
                    crate::vbscript::expr::BinOp::IntDiv => Instruction::IntDiv,
                    crate::vbscript::expr::BinOp::Mod => Instruction::Mod,
                    crate::vbscript::expr::BinOp::Pow => Instruction::Pow,
                    crate::vbscript::expr::BinOp::Concat => Instruction::Concat,
                    crate::vbscript::expr::BinOp::Eq => Instruction::Eq,
                    crate::vbscript::expr::BinOp::Ne => Instruction::Ne,
                    crate::vbscript::expr::BinOp::Lt => Instruction::Lt,
                    crate::vbscript::expr::BinOp::Gt => Instruction::Gt,
                    crate::vbscript::expr::BinOp::Le => Instruction::Le,
                    crate::vbscript::expr::BinOp::Ge => Instruction::Ge,
                    crate::vbscript::expr::BinOp::And => Instruction::And,
                    crate::vbscript::expr::BinOp::Or => Instruction::Or,
                    crate::vbscript::expr::BinOp::Xor => Instruction::Xor,
                    crate::vbscript::expr::BinOp::Eqv => Instruction::Eqv,
                    crate::vbscript::expr::BinOp::Imp => Instruction::Imp,
                    crate::vbscript::expr::BinOp::Is => Instruction::Is,
                    crate::vbscript::expr::BinOp::Like => Instruction::Like,
                };
                self.emit(inst);
            }
            Expr::UnaryOp { op, expr: inner } => {
                self.compile_expr(inner);
                let inst = match op {
                    crate::vbscript::expr::UnaryOp::Neg => Instruction::Neg,
                    crate::vbscript::expr::UnaryOp::Not => Instruction::Not,
                };
                self.emit(inst);
            }
            Expr::FunctionCall { name, args } => {
                let name_lower = name.to_lowercase();
                // Check if name matches a known local variable -> array read
                if let Some(slot) = self.local_slot(&name_lower) {
                    if args.len() > 1 {
                        // Multi-dimensional array access
                        for arg in args {
                            self.compile_expr(arg);
                        }
                        self.emit(Instruction::CallLocal(slot, args.len() as u8));
                    } else {
                        self.emit(Instruction::LoadLocal(slot));
                        for arg in args {
                            self.compile_expr(arg);
                        }
                        self.emit(Instruction::IndexGet);
                    }
                } else {
                    for arg in args {
                        self.compile_expr(arg);
                    }
                    let name_idx = self.add_constant(VBValue::String(name_lower.into()));
                    self.emit(Instruction::Call(name_idx, args.len() as u8));
                }
            }
            Expr::PropertyAccess { object, property } => {
                self.compile_expr(object);
                let prop_idx = self.add_constant(VBValue::String(property.to_lowercase().into()));
                self.emit(Instruction::GetProp(prop_idx));
            }
            Expr::MethodCall {
                object,
                method,
                args,
            } => {
                self.compile_expr(object);
                for arg in args {
                    self.compile_expr(arg);
                }
                let method_idx = self.add_constant(VBValue::String(method.to_lowercase().into()));
                self.emit(Instruction::CallMethod(method_idx, args.len() as u8));
            }
            Expr::NewObject(name) => {
                let name_idx = self.add_constant(VBValue::String(name.to_lowercase().into()));
                self.emit(Instruction::NewObject(name_idx));
            }
            Expr::WithObject => {
                let name_idx = self.add_constant(VBValue::String("__with_obj__".into()));
                self.emit(Instruction::LoadGlobal(name_idx));
            }
            Expr::CaseComparison { op, rhs } => {
                self.emit(Instruction::LoadSelectValue);
                self.compile_expr(rhs);
                let inst = match op {
                    crate::vbscript::expr::BinOp::Eq => Instruction::Eq,
                    crate::vbscript::expr::BinOp::Ne => Instruction::Ne,
                    crate::vbscript::expr::BinOp::Lt => Instruction::Lt,
                    crate::vbscript::expr::BinOp::Le => Instruction::Le,
                    crate::vbscript::expr::BinOp::Gt => Instruction::Gt,
                    crate::vbscript::expr::BinOp::Ge => Instruction::Ge,
                    _ => Instruction::Eq,
                };
                self.emit(inst);
            }
            Expr::Range { low, high } => {
                self.emit(Instruction::LoadSelectValue);
                self.compile_expr(low);
                self.emit(Instruction::Ge);
                self.emit(Instruction::LoadSelectValue);
                self.compile_expr(high);
                self.emit(Instruction::Le);
                self.emit(Instruction::And);
            }
        }
    }
}

pub(crate) fn extract_properties_from_class_body(
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

pub(crate) fn extract_methods_from_class_body(body_lines: &[Vec<Token>]) -> AHashMap<String, MethodDef> {
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
