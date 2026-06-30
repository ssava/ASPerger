use std::sync::Arc;
use crate::vbscript::builtins;
use crate::vbscript::compiler::CompiledCode;
use crate::vbscript::execution_context::ErrorMode;
use crate::vbscript::instruction::Instruction;
use crate::vbscript::value_utils;
use crate::vbscript::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::{ExecutionContext, VBValue};

pub struct Vm<'a> {
    code: Arc<Vec<Instruction>>,
    constants: Arc<Vec<VBValue>>,
    ip: usize,
    stack: Vec<VBValue>,
    locals: Vec<VBValue>,
    frames: Vec<CallFrame>,
    for_states: Vec<ForState>,
    for_each_states: Vec<ForEachState>,
    select_value: Option<VBValue>,
    with_stack: Vec<VBValue>,
    context: &'a mut ExecutionContext,
    should_exit: bool,
}

#[allow(dead_code)]
struct CallFrame {
    return_ip: usize,
    code: Arc<Vec<Instruction>>,
    constants: Arc<Vec<VBValue>>,
    stack_base: usize,
    locals_count: usize,
}

struct ForState {
    counter_slot: usize,
    end: VBValue,
    step: VBValue,
}

struct ForEachState {
    element_slot: usize,
    array: Arc<Vec<VBValue>>,
    index: usize,
}

impl<'a> Vm<'a> {
    pub fn new(context: &'a mut ExecutionContext) -> Self {
        Vm {
            code: Arc::new(Vec::new()),
            constants: Arc::new(Vec::new()),
            ip: 0,
            stack: Vec::new(),
            locals: Vec::new(),
            frames: Vec::new(),
            for_states: Vec::new(),
            for_each_states: Vec::new(),
            select_value: None,
            with_stack: Vec::new(),
            context,
            should_exit: false,
        }
    }

    pub fn run(&mut self, compiled: CompiledCode) -> Result<(), VBSError> {
        self.code = Arc::new(compiled.instructions);
        self.constants = Arc::new(compiled.constants);
        self.locals = vec![VBValue::Empty; compiled.local_count];
        self.ip = 0;
        self.stack.clear();
        self.frames.clear();
        self.for_states.clear();
        self.for_each_states.clear();
        self.select_value = None;
        self.with_stack.clear();
        self.should_exit = false;

        self.execute_loop()
    }

    fn execute_loop(&mut self) -> Result<(), VBSError> {
        let mut iterations = 0u64;
        loop {
            if self.should_exit {
                return Ok(());
            }
            iterations += 1;
            if iterations > 10_000_000 {
                return Err(VBSError::new(0, format!("VM iteration limit exceeded at ip={}", self.ip), VBSErrorType::RuntimeError));
            }
            if self.context.response.ended {
                return Ok(());
            }
            if self.ip >= self.code.len() {
                return Ok(());
            }
            let inst = self.code[self.ip].clone();
            self.ip += 1;

            match inst {
                // -- Constants --
                Instruction::LoadConst(i) => {
                    let val = self.constants[i as usize].clone();
                    self.stack.push(val);
                }
                Instruction::LoadNil => self.stack.push(VBValue::Null),
                Instruction::LoadTrue => self.stack.push(VBValue::Boolean(true)),
                Instruction::LoadFalse => self.stack.push(VBValue::Boolean(false)),
                Instruction::LoadEmpty => self.stack.push(VBValue::Empty),

                // -- Variables --
                Instruction::LoadLocal(s) => {
                    self.stack.push(self.locals[s as usize].clone());
                }
                Instruction::StoreLocal(s) => {
                    let val = self.stack.pop().unwrap();
                    self.locals[s as usize] = val;
                }
                Instruction::LoadGlobal(i) => {
                    let name = self.constants[i as usize].to_string();
                    let val = self
                        .context
                        .get_variable(&name)
                        .cloned()
                        .unwrap_or(VBValue::Empty);
                    self.stack.push(val);
                }
                Instruction::StoreGlobal(i) => {
                    let name = self.constants[i as usize].to_string();
                    let val = self.stack.pop().unwrap();
                    self.context.set_variable(&name, val);
                }

                // -- Unary --
                Instruction::Neg => {
                    let val = self.stack.pop().unwrap();
                    self.stack.push(Vm::negate(val));
                }
                Instruction::Not => {
                    let val = self.stack.pop().unwrap();
                    self.stack.push(Vm::logical_not(val));
                }

                // -- Binary arithmetic --
                Instruction::Add => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::add(l, r));
                }
                Instruction::Sub => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::sub(l, r));
                }
                Instruction::Mul => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::mul(l, r));
                }
                Instruction::Div => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::div(l, r));
                }
                Instruction::IntDiv => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::int_div(l, r));
                }
                Instruction::Mod => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::mod_op(l, r));
                }
                Instruction::Pow => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::pow_op(l, r));
                }

                // -- String --
                Instruction::Concat => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::concat_str(l, r));
                }

                // -- Comparison --
                Instruction::Eq => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(VBValue::Boolean(Vm::values_equal(&l, &r)));
                }
                Instruction::Ne => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(VBValue::Boolean(!Vm::values_equal(&l, &r)));
                }
                Instruction::Lt => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::compare_lt(l, r));
                }
                Instruction::Le => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::compare_le(l, r));
                }
                Instruction::Gt => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::compare_gt(l, r));
                }
                Instruction::Ge => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::compare_ge(l, r));
                }
                Instruction::Is => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(VBValue::Boolean(Vm::values_equal(&l, &r)));
                }
                Instruction::Like => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::like_match(l, r));
                }

                // -- Logical --
                Instruction::And => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::bool_or_bitwise(l, r, |a, b| a & b));
                }
                Instruction::Or => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::bool_or_bitwise(l, r, |a, b| a | b));
                }
                Instruction::Xor => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::bool_or_bitwise(l, r, |a, b| a ^ b));
                }
                Instruction::Imp => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::imp_op(l, r));
                }
                Instruction::Eqv => {
                    let r = self.stack.pop().unwrap();
                    let l = self.stack.pop().unwrap();
                    self.stack.push(Vm::eqv_op(l, r));
                }

                // -- Objects --
                Instruction::GetProp(i) => {
                    let obj = self.stack.pop().unwrap();
                    let prop = self.constants[i as usize].to_string();
                    match obj {
                        VBValue::Object(obj) => {
                            let result = match obj.get_property(&prop, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Property '{}' not found: {}", prop, e),
                                    VBSErrorType::RuntimeError
                                )) {
                                Ok(v) => v,
                                Err(e) => {
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            };
                            self.stack.push(result);
                        }
                        _ => {
                            let e = VBSError::new(0, format!("Object required for property access: {}", prop), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::SetProp(i) => {
                    let val = self.stack.pop().unwrap();
                    let mut obj = self.stack.pop().unwrap();
                    let prop = self.constants[i as usize].to_string();
                    match &mut obj {
                        VBValue::Object(obj) => {
                            match obj.set_property(&prop, val, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Cannot set property '{}': {}", prop, e),
                                    VBSErrorType::RuntimeError
                                )) {
                                Ok(_) => {}
                                Err(e) => {
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            }
                        }
                        _ => {
                            let e = VBSError::new(0, "Object required".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::CallMethod(i, n) => {
                    let method = self.constants[i as usize].to_string();
                    let n_args = n as usize;

                    let args: Vec<VBValue> = if n_args > 0 {
                        let start = self.stack.len() - n_args;
                        self.stack.drain(start..).collect()
                    } else {
                        Vec::new()
                    };

                    let mut obj = self.stack.pop().unwrap();
                    match &mut obj {
                        VBValue::Object(obj) => {
                            // First try property + indexed access pattern
                            let found = if n_args == 1 && !args.is_empty() {
                                if let Ok(sub) = obj.get_property(&method, self.context) {
                                    if let VBValue::Object(sub_obj) = sub {
                                        if let Ok(result) = sub_obj.indexed_get(&args[0], self.context) {
                                            self.stack.push(result);
                                            true
                                        } else { false }
                                    } else { false }
                                } else { false }
                            } else { false };

                            if !found {
                                let result = match obj.call_method(&method, &args, self.context)
                                    .map_err(|e| VBSError::new(
                                        0, format!("Method '{}' failed: {}", method, e),
                                        VBSErrorType::RuntimeError
                                    )) {
                                    Ok(v) => v,
                                    Err(e) => {
                                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                            self.context.set_err(e);
                                            continue;
                                        } else {
                                            return Err(e);
                                        }
                                    }
                                };
                                self.stack.push(result);
                            }
                        }
                        _ => {
                            let e = VBSError::new(0, format!("Object required for method call: {}", method), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::CallMethodLocal(slot, i, n) => {
                    let method = self.constants[i as usize].to_string();
                    let n_args = n as usize;

                    let args: Vec<VBValue> = if n_args > 0 {
                        let start = self.stack.len() - n_args;
                        self.stack.drain(start..).collect()
                    } else {
                        Vec::new()
                    };

                    let mut obj_val = VBValue::Empty;
                    std::mem::swap(&mut self.locals[slot as usize], &mut obj_val);
                    let result = match &mut obj_val {
                        VBValue::Object(obj) => {
                            obj.call_method(&method, &args, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Method '{}' failed: {}", method, e),
                                    VBSErrorType::RuntimeError
                                ))
                        }
                        _ => {
                            Err(VBSError::new(
                                0, format!("Object required for method call: {}", method),
                                VBSErrorType::RuntimeError
                            ))
                        }
                    };
                    std::mem::swap(&mut self.locals[slot as usize], &mut obj_val);
                    match result {
                        Ok(v) => self.stack.push(v),
                        Err(e) => {
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::CallMethodGlobal(g, i, n) => {
                    let method = self.constants[i as usize].to_string();
                    let name = self.constants[g as usize].to_string();
                    let n_args = n as usize;

                    let args: Vec<VBValue> = if n_args > 0 {
                        let start = self.stack.len() - n_args;
                        self.stack.drain(start..).collect()
                    } else {
                        Vec::new()
                    };

                    let mut obj_val = VBValue::Empty;
                    if let Some(slot) = self.context.get_variable_mut(&name) {
                        std::mem::swap(slot, &mut obj_val);
                    }
                    let result = match &mut obj_val {
                        VBValue::Object(obj) => {
                            obj.call_method(&method, &args, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Method '{}' failed: {}", method, e),
                                    VBSErrorType::RuntimeError
                                ))
                        }
                        _ => {
                            Err(VBSError::new(
                                0, format!("Object required for method call: {}", method),
                                VBSErrorType::RuntimeError
                            ))
                        }
                    };
                    self.context.set_variable(&name, obj_val);
                    match result {
                        Ok(v) => self.stack.push(v),
                        Err(e) => {
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::IndexGet => {
                    let key = self.stack.pop().unwrap();
                    let obj = self.stack.pop().unwrap();
                    match obj {
                        VBValue::Object(obj) => {
                            let result = match obj.indexed_get(&key, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Indexed get failed: {}", e),
                                    VBSErrorType::RuntimeError
                                )) {
                                Ok(v) => v,
                                Err(e) => {
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            };
                            self.stack.push(result);
                        }
                        VBValue::Array(arr, _dims) => {
                            let idx = value_utils::to_arg_f64(&key) as usize;
                            if idx < arr.len() {
                                self.stack.push(arr[idx].clone());
                            } else {
                                let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                        }
                        _ => {
                            let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::IndexSet => {
                    let val = self.stack.pop().unwrap();
                    let key = self.stack.pop().unwrap();
                    let mut obj = self.stack.pop().unwrap();
                    match &mut obj {
                        VBValue::Object(obj) => {
                            match obj.indexed_set(&key, val, self.context)
                                .map_err(|e| VBSError::new(
                                    0, format!("Indexed set failed: {}", e),
                                    VBSErrorType::RuntimeError
                                )) {
                                Ok(_) => {}
                                Err(e) => {
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            }
                        }
                        VBValue::Array(arr_ref, _) => {
                            let idx = value_utils::to_arg_f64(&key) as usize;
                            if let Some(arr) = Arc::get_mut(arr_ref) {
                                if idx < arr.len() {
                                    arr[idx] = val;
                                } else {
                                    let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                    } else {
                                        return Err(e);
                                    }
                                }
                            } else {
                                let e = VBSError::new(0, "Cannot mutate shared array".to_string(), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                        }
                        _ => {
                            let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::IndexStoreLocal(slot) => {
                    let val = self.stack.pop().unwrap();
                    let key = self.stack.pop().unwrap();
                    let arr = &mut self.locals[slot as usize];
                    if let VBValue::Array(arr_ref, _) = arr {
                        let idx = value_utils::to_arg_f64(&key) as usize;
                        let items = Arc::make_mut(arr_ref);
                        if idx < items.len() {
                            items[idx] = val;
                        } else {
                            let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    } else {
                        let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                            self.context.set_err(e);
                        } else {
                            return Err(e);
                        }
                    }
                }
                Instruction::IndexStoreGlobal(idx) => {
                    let val = self.stack.pop().unwrap();
                    let key = self.stack.pop().unwrap();
                    let name = self.constants[idx as usize].to_string();
                    let var = self.context.get_variable_mut(&name);
                    if let Some(VBValue::Array(arr_ref, _)) = var {
                        let idx_val = value_utils::to_arg_f64(&key) as usize;
                        let items = Arc::make_mut(arr_ref);
                        if idx_val < items.len() {
                            items[idx_val] = val;
                        } else {
                            let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    } else {
                        let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                            self.context.set_err(e);
                        } else {
                            return Err(e);
                        }
                    }
                }
                Instruction::IndexStoreLocalMulti(slot, n) => {
                    let n_indices = n as usize;
                    let val = self.stack.pop().unwrap();
                    let start = self.stack.len() - n_indices;
                    let indices: Vec<VBValue> = self.stack.drain(start..).collect();
                    let arr = &mut self.locals[slot as usize];
                    if let VBValue::Array(arr_ref, dims) = arr {
                        let flat_idx = if dims.is_empty() && n_indices == 1 {
                            let idx_val = value_utils::to_arg_f64(&indices[0]) as usize;
                            if idx_val >= arr_ref.len() {
                                let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                            idx_val
                        } else if n_indices == dims.len() {
                            let flat_idx = match value_utils::compute_flat_index(&indices, dims) {
                                Some(v) => v,
                                None => {
                                    let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            };
                            flat_idx
                        } else {
                            let e = VBSError::new(9, format!("Array has {} dimensions but {} indices provided", dims.len(), n_indices), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                                continue;
                            } else {
                                return Err(e);
                            }
                        };
                        let items = Arc::make_mut(arr_ref);
                        if flat_idx < items.len() {
                            items[flat_idx] = val;
                        } else {
                            let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    } else {
                        let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                            self.context.set_err(e);
                            continue;
                        } else {
                            return Err(e);
                        }
                    }
                }
                Instruction::IndexStoreGlobalMulti(idx, n) => {
                    let n_indices = n as usize;
                    let val = self.stack.pop().unwrap();
                    let start = self.stack.len() - n_indices;
                    let indices: Vec<VBValue> = self.stack.drain(start..).collect();
                    let name = self.constants[idx as usize].to_string();
                    let var = self.context.get_variable_mut(&name);
                    if let Some(VBValue::Array(arr_ref, dims)) = var {
                        let flat_idx = if dims.is_empty() && n_indices == 1 {
                            let idx_val = value_utils::to_arg_f64(&indices[0]) as usize;
                            if idx_val >= arr_ref.len() {
                                let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                    continue;
                                } else {
                                    return Err(e);
                                }
                            }
                            idx_val
                        } else if n_indices == dims.len() {
                            let flat_idx = match value_utils::compute_flat_index(&indices, dims) {
                                Some(v) => v,
                                None => {
                                    let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            };
                            flat_idx
                        } else {
                            let e = VBSError::new(9, format!("Array has {} dimensions but {} indices provided", dims.len(), n_indices), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                                continue;
                            } else {
                                return Err(e);
                            }
                        };
                        let items = Arc::make_mut(arr_ref);
                        if flat_idx < items.len() {
                            items[flat_idx] = val;
                        } else {
                            let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    } else {
                        let e = VBSError::new(0, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                            self.context.set_err(e);
                        } else {
                            return Err(e);
                        }
                    }
                }
                Instruction::NewObject(i) => {
                    let name = self.constants[i as usize].to_string();
                    let class = self.context.get_class(&name);
                    if class.is_some() {
                        let instance = crate::vbscript::vbobject::ClassInstance::new(&name);
                        self.stack.push(VBValue::Object(Box::new(instance)));
                    } else {
                        let e = VBSError::new(0, format!("Class '{}' not found", name), VBSErrorType::RuntimeError);
                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                            self.context.set_err(e);
                        } else {
                            return Err(e);
                        }
                    }
                }

                // -- Arrays --
                Instruction::NewArray(n) => {
                    let dim_count = n as usize;
                    let mut dims = Vec::with_capacity(dim_count);
                    let mut total_size = 1;
                    for _ in 0..dim_count {
                        let size = value_utils::to_arg_f64(&self.stack.pop().unwrap()) as usize;
                        dims.push(size);
                        total_size *= size;
                    }
                    dims.reverse();
                    let arr = vec![VBValue::Empty; total_size];
                    self.stack.push(VBValue::Array(Arc::new(arr), dims));
                }
                Instruction::ReDim(slot, n, preserve) => {
                    let dim_count = n as usize;
                    let mut dims = Vec::with_capacity(dim_count);
                    let mut total_size = 1;
                    let mut new_dims = Vec::with_capacity(dim_count);
                    for _ in 0..dim_count {
                        let bound = value_utils::to_arg_f64(&self.stack.pop().unwrap()) as usize;
                        dims.push(bound);
                        new_dims.push(bound);
                        total_size *= bound + 1;
                    }
                    dims.reverse();
                    new_dims.reverse();

                    if preserve {
                        if let VBValue::Array(old_arr, _old_dims) = &self.locals[slot as usize] {
                            let mut new_arr = vec![VBValue::Empty; total_size];
                            let copy_len = old_arr.len().min(total_size);
                            for i in 0..copy_len {
                                new_arr[i] = old_arr[i].clone();
                            }
                            self.locals[slot as usize] = VBValue::Array(Arc::new(new_arr), new_dims);
                        } else {
                            self.locals[slot as usize] = VBValue::Array(Arc::new(vec![VBValue::Empty; total_size]), new_dims);
                        }
                    } else {
                        self.locals[slot as usize] = VBValue::Array(Arc::new(vec![VBValue::Empty; total_size]), new_dims);
                    }
                }

                // -- Functions --
                Instruction::Call(i, n) => {
                    let name = self.constants[i as usize].to_string();
                    let n_args = n as usize;

                    let args: Vec<VBValue> = if n_args > 0 {
                        let start = self.stack.len() - n_args;
                        self.stack.drain(start..).collect()
                    } else {
                        Vec::new()
                    };

                    // Check array access first (matches old interpreter's evaluate order)
                    if n_args > 0 {
                        if let Some(VBValue::Array(items, dims)) = self.context.get_variable(&name) {
                            let flat_idx = if dims.is_empty() && args.len() == 1 {
                                let idx = value_utils::to_arg_f64(&args[0]) as usize;
                                if idx >= items.len() {
                                    let e = VBSError::new(9, format!("Subscript out of range: index {} exceeds array size {}", idx, items.len()), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                                idx
                            } else if args.len() == dims.len() {
                                let idx = match value_utils::compute_flat_index(&args, dims) {
                                    Some(v) => v,
                                    None => {
                                        let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                            self.context.set_err(e);
                                            continue;
                                        } else {
                                            return Err(e);
                                        }
                                    }
                                };
                                if idx >= items.len() {
                                    let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                                idx
                            } else {
                                let e = VBSError::new(9, format!("Array has {} dimensions but {} indices provided", dims.len(), args.len()), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                    continue;
                                } else {
                                    return Err(e);
                                }
                            };
                            self.stack.push(items[flat_idx].clone());
                            continue;
                        }
                    }

                    // Check object indexed access (e.g. dict("key"))
                    if n_args > 0 && self.context.get_variable(&name).is_some() {
                        let is_obj = matches!(self.context.get_variable(&name), Some(VBValue::Object(_)));
                        if is_obj {
                            let key = &args[0];
                            let mut obj_val = VBValue::Empty;
                            if let Some(slot) = self.context.get_variable_mut(&name) {
                                std::mem::swap(slot, &mut obj_val);
                            }
                            let result = match &mut obj_val {
                                VBValue::Object(ref mut obj) => {
                                    obj.indexed_get(key, self.context)
                                }
                                _ => unreachable!(),
                            };
                            self.context.set_variable(&name, obj_val);
                            match result {
                                Ok(val) => {
                                    self.stack.push(val);
                                    continue;
                                }
                                Err(e) => {
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                            }
                        }
                    }

                    // Check user-defined function
                    if self.context.get_function(&name).is_some() {
                        match self.call_user_function(&name, &args) {
                            Ok(()) => {}
                            Err(e) => {
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                        }
                    } else {
                        // Check built-in
                        match builtins::call_builtin(&name, args) {
                            Ok(v) => self.stack.push(v),
                            Err(e) => {
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                        }
                    }
                }
                Instruction::CallLocal(slot, n) => {
                    let n_args = n as usize;
                    let args: Vec<VBValue> = if n_args > 0 {
                        let start = self.stack.len() - n_args;
                        self.stack.drain(start..).collect()
                    } else {
                        Vec::new()
                    };
                    // Load the local variable
                    let local_val = self.locals[slot].clone();
                    match local_val {
                        VBValue::Array(items, dims) => {
                            let flat_idx = if dims.is_empty() && args.len() == 1 {
                                let idx = value_utils::to_arg_f64(&args[0]) as usize;
                                if idx >= items.len() {
                                    let e = VBSError::new(9, format!("Subscript out of range: index {} exceeds array size {}", idx, items.len()), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                                idx
                            } else if args.len() == dims.len() {
                                let idx = match value_utils::compute_flat_index(&args, &dims) {
                                    Some(v) => v,
                                    None => {
                                        let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                            self.context.set_err(e);
                                            continue;
                                        } else {
                                            return Err(e);
                                        }
                                    }
                                };
                                if idx >= items.len() {
                                    let e = VBSError::new(9, "Subscript out of range".to_string(), VBSErrorType::RuntimeError);
                                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                        self.context.set_err(e);
                                        continue;
                                    } else {
                                        return Err(e);
                                    }
                                }
                                idx
                            } else {
                                let e = VBSError::new(9, format!("Array has {} dimensions but {} indices provided", dims.len(), args.len()), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                    continue;
                                } else {
                                    return Err(e);
                                }
                            };
                            self.stack.push(items[flat_idx].clone());
                        }
                        VBValue::Object(obj) => {
                            if let Some(arg) = args.first() {
                                match obj.indexed_get(arg, self.context) {
                                    Ok(val) => self.stack.push(val),
                                    Err(e) => {
                                        if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                            self.context.set_err(e);
                                        } else {
                                            return Err(e);
                                        }
                                    }
                                }
                            } else {
                                self.stack.push(VBValue::Empty);
                            }
                        }
                        _ => {
                            let e = VBSError::new(424, "Object or array required".to_string(), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                                continue;
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::Return(n) => {
                    let frame = self.frames.pop().unwrap();
                    let result = if n > 0 {
                        self.stack.pop().unwrap_or(VBValue::Empty)
                    } else {
                        VBValue::Empty
                    };
                    self.stack.truncate(frame.stack_base);
                    self.code = frame.code;
                    self.constants = frame.constants;
                    self.ip = frame.return_ip;
                    if n > 0 {
                        self.stack.push(result);
                    }
                    if self.frames.is_empty() && self.ip >= self.code.len() {
                        return Ok(());
                    }
                }

                // -- Control flow --
                Instruction::Jump(offset) => {
                    self.ip = (self.ip as isize + offset as isize) as usize;
                }
                Instruction::JumpIfFalse(offset) => {
                    let val = self.stack.pop().unwrap();
                    if !Vm::is_truthy(&val) {
                        self.ip = (self.ip as isize + offset as isize) as usize;
                    }
                }
                Instruction::JumpIfTrue(offset) => {
                    let val = self.stack.pop().unwrap();
                    if Vm::is_truthy(&val) {
                        self.ip = (self.ip as isize + offset as isize) as usize;
                    }
                }

                // -- Loops --
                Instruction::ForPrep(slot, exit_offset) => {
                    let existing = self.for_states.iter().position(|fs| fs.counter_slot == slot);
                    if let Some(pos) = existing {
                        let end = self.for_states[pos].end.clone();
                        let step = self.for_states[pos].step.clone();
                        if Vm::is_past_end(&self.locals[slot as usize], &end, &step) {
                            self.for_states.remove(pos);
                            self.ip = (self.ip as isize + exit_offset as isize) as usize;
                        }
                    } else {
                        let step = self.stack.pop().unwrap();
                        let end = self.stack.pop().unwrap();
                        let counter = self.locals[slot as usize].clone();
                        self.for_states.push(ForState {
                            counter_slot: slot,
                            end,
                            step: step.clone(),
                        });
                        if Vm::is_past_end(&counter, &self.for_states.last().unwrap().end, &step) {
                            self.for_states.pop();
                            self.ip = (self.ip as isize + exit_offset as isize) as usize;
                        }
                    }
                }
                Instruction::ForStep(slot, back_offset) => {
                    if let Some(fs) = self.for_states.last() {
                        if fs.counter_slot == slot {
                            let step_val = value_utils::to_arg_f64(&fs.step);
                            match &self.locals[slot as usize] {
                                VBValue::Number(n) => {
                                    self.locals[slot as usize] = VBValue::Number(n + step_val);
                                }
                                _ => {
                                    self.locals[slot as usize] = VBValue::Number(step_val);
                                }
                            }
                            if !Vm::is_past_end(&self.locals[slot as usize], &fs.end, &fs.step) {
                                self.ip = (self.ip as isize + back_offset as isize) as usize;
                            } else {
                                self.for_states.pop();
                            }
                        } else {
                            self.for_states.pop();
                        }
                    }
                }
                Instruction::ForEachPrep(slot, exit_offset) => {
                    let existing = self.for_each_states.iter().position(|fes| fes.element_slot == slot);
                    if existing.is_some() {
                        // Subsequent iteration — ForEachStep already updated the slot
                    } else {
                        let group = self.stack.pop().unwrap();
                        match group {
                            VBValue::Array(arr, _) => {
                                if arr.is_empty() {
                                    self.ip = (self.ip as isize + exit_offset as isize) as usize;
                                } else {
                                    self.locals[slot as usize] = arr[0].clone();
                                    self.for_each_states.push(ForEachState {
                                        element_slot: slot,
                                        array: arr,
                                        index: 0,
                                    });
                                }
                            }
                            VBValue::Object(obj) => {
                                if let Ok(VBValue::Array(keys_arr, _)) = obj.get_property("keys", self.context) {
                                    if keys_arr.is_empty() {
                                        self.ip = (self.ip as isize + exit_offset as isize) as usize;
                                    } else {
                                        self.locals[slot as usize] = keys_arr[0].clone();
                                        self.for_each_states.push(ForEachState {
                                            element_slot: slot,
                                            array: keys_arr,
                                            index: 0,
                                        });
                                    }
                                } else {
                                    self.ip = (self.ip as isize + exit_offset as isize) as usize;
                                }
                            }
                            _ => {
                                let e = VBSError::new(0, "For Each requires an array or object".to_string(), VBSErrorType::RuntimeError);
                                if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                    self.context.set_err(e);
                                } else {
                                    return Err(e);
                                }
                            }
                        }
                    }
                }
                Instruction::ForEachStep(slot, back_offset) => {
                    if let Some(fes) = self.for_each_states.last_mut() {
                        if fes.element_slot == slot {
                            fes.index += 1;
                            if fes.index < fes.array.len() {
                                self.locals[slot as usize] = fes.array[fes.index].clone();
                                self.ip = (self.ip as isize + back_offset as isize) as usize;
                            } else {
                                self.for_each_states.pop();
                            }
                        } else {
                            self.for_each_states.pop();
                        }
                    }
                }

                // -- Exit signals --
                Instruction::ExitFor |
                Instruction::ExitDo => {
                    // Jump past the loop. Find the next instruction after the loop
                    // by scanning the ForStep/ForEachStep/back-jump.
                    // For now, just return.
                    // We'll handle this via the loop_stack in the compiler
                }
                Instruction::ExitFunction => {
                    self.should_exit = true;
                }
                Instruction::ExitSub => {
                    self.should_exit = true;
                }

                // -- Scope --
                Instruction::SelectStore => {
                    self.select_value = self.stack.pop();
                }
                Instruction::LoadSelectValue => {
                    if let Some(ref sel) = self.select_value {
                        self.stack.push(sel.clone());
                    } else {
                        self.stack.push(VBValue::Empty);
                    }
                }
                Instruction::SelectCompare => {
                    let val = self.stack.pop().unwrap();
                    if let Some(ref sel) = self.select_value {
                        self.stack.push(VBValue::Boolean(Vm::values_equal(sel, &val)));
                    } else {
                        self.stack.push(VBValue::Boolean(false));
                    }
                }
                Instruction::SelectClear => {
                    self.select_value = None;
                }
                Instruction::WithStart => {
                    let obj = self.stack.pop().unwrap();
                    self.with_stack.push(obj.clone());
                    self.context.set_variable("__with_obj__", obj);
                }
                Instruction::WithEnd => {
                    self.with_stack.pop();
                    if let Some(prev) = self.with_stack.last() {
                        self.context.set_variable("__with_obj__", prev.clone());
                    } else {
                        self.context.set_variable("__with_obj__", VBValue::Empty);
                    }
                }

                // -- Error handling --
                Instruction::OnErrorResumeNext => {
                    self.context.set_error_mode(ErrorMode::ResumeNext);
                }
                Instruction::OnErrorGoto0 => {
                    self.context.set_error_mode(ErrorMode::Normal);
                }
                Instruction::Raise(i) => {
                    let _data = &self.constants[i as usize];
                    let e = VBSError::new(0, "Error raised".to_string(), VBSErrorType::RuntimeError);
                    if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                        self.context.set_err(e);
                    } else {
                        return Err(e);
                    }
                }

                // -- Debug --
                Instruction::DebugLine(_line) => {
                    // No-op when debugger is not attached
                }

                // -- ASP-specific --
                Instruction::ResponseWrite => {
                    let val = self.stack.pop().unwrap();
                    self.context.write(&val.to_string());
                }
                Instruction::ResponseEnd => {
                    self.context.response.ended = true;
                    return Ok(());
                }
                Instruction::ServerExecute(i) => {
                    let path = self.constants[i as usize].to_string();
                    let cb = self.context.execute_file_callback.clone();
                    if let Some(cb) = cb {
                        if let Err(e) = cb(&path, self.context) {
                            let e = VBSError::new(0, format!("Server.Execute failed: {}", e), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }
                Instruction::ServerTransfer(i) => {
                    let path = self.constants[i as usize].to_string();
                    let cb = self.context.execute_file_callback.clone();
                    if let Some(cb) = cb {
                        if let Err(e) = cb(&path, self.context) {
                            let e = VBSError::new(0, format!("Server.Transfer failed: {}", e), VBSErrorType::RuntimeError);
                            if *self.context.get_error_mode() == ErrorMode::ResumeNext {
                                self.context.set_err(e);
                            } else {
                                return Err(e);
                            }
                        }
                    }
                }

                // -- Variable management --
                Instruction::Erase(slot) => {
                    self.locals[slot as usize] = VBValue::Empty;
                }
            }
        }
    }

    fn call_user_function(&mut self, name: &str, args: &[VBValue]) -> Result<(), VBSError> {
        let func = self.context.get_function(name)
            .ok_or_else(|| VBSError::new(
                0, format!("Function '{}' not found", name),
                VBSErrorType::RuntimeError
            ))?
            .clone();

        let func_name = func.name.clone();
        let is_func = func.is_function;

        // Save/reset code_start_line
        let saved_code_start_line = self.context.code_start_line;
        self.context.code_start_line = 0;

        // Set up parameters in context for global access (e.g. Return statement stores function name as global)
        for (i, param) in func.params.iter().enumerate() {
            let val = args.get(i).cloned().unwrap_or(VBValue::Empty);
            self.context.set_variable(param, val);
        }
        if is_func {
            self.context.set_variable(&func_name, VBValue::Empty);
        }

        // Get or compile the function code
        let func_code = if let Some(cached) = self.context.get_function_code(&func_name) {
            cached.clone()
        } else {
            let lines = &func.body_lines;
            let blocks = crate::vbscript::block::parse_blocks(lines)
                .map_err(|e| VBSError::new(
                    0, format!("Failed to parse function '{}': {}", name, e),
                    VBSErrorType::RuntimeError
                ))?;
            let mut compiler = crate::vbscript::compiler::Compiler::new(self.context);
            let compiled = compiler.compile(&blocks)
                .map_err(|e| VBSError::new(
                    0, format!("Failed to compile function '{}': {}", name, e),
                    VBSErrorType::RuntimeError
                ))?;
            for (n, c) in compiled.compiled_functions.clone() {
                self.context.set_function_code(&n, c);
            }
            self.context.set_function_code(&func_name, compiled.clone());
            compiled
        };

        // Save outer VM state
        let saved_ip = self.ip;
        let saved_code = std::mem::replace(&mut self.code, Arc::new(func_code.instructions));
        let saved_constants = std::mem::replace(&mut self.constants, Arc::new(func_code.constants));
        let saved_locals = std::mem::replace(&mut self.locals, vec![VBValue::Empty; func_code.local_count]);
        let saved_stack = std::mem::take(&mut self.stack);
        let saved_for_states = std::mem::take(&mut self.for_states);
        let saved_for_each_states = std::mem::take(&mut self.for_each_states);
        let saved_select_value = self.select_value.take();
        let saved_with_stack = std::mem::take(&mut self.with_stack);
        let saved_frames = std::mem::take(&mut self.frames);

        let saved_should_exit = self.should_exit;
        self.should_exit = false;
        self.ip = 0;

        // Set parameters in locals (already initialized to Empty above)
        for (i, _param) in func.params.iter().enumerate() {
            if i < args.len() {
                self.locals[i] = args[i].clone();
            }
        }

        // Execute function body via VM
        let result = self.execute_loop();

        // Extract return value
        let return_val = if is_func {
            self.context.get_variable(&func_name).cloned().unwrap_or(VBValue::Empty)
        } else {
            VBValue::Empty
        };

        // Restore outer VM state entirely
        self.should_exit = saved_should_exit;
        self.code = saved_code;
        self.constants = saved_constants;
        self.locals = saved_locals;
        self.stack = saved_stack;
        self.for_states = saved_for_states;
        self.for_each_states = saved_for_each_states;
        self.select_value = saved_select_value;
        self.with_stack = saved_with_stack;
        self.frames = saved_frames;
        self.ip = saved_ip;

        self.context.code_start_line = saved_code_start_line;

        if is_func {
            self.stack.push(return_val);
        }

        match result {
            Ok(()) => Ok(()),
            Err(e) if e.is_exit_function() || e.is_exit_sub() => Ok(()),
            Err(e) => Err(e),
        }
    }

    // -- Helper functions (ported from the existing interpreter) --

    fn is_truthy(val: &VBValue) -> bool {
        match val {
            VBValue::Boolean(b) => *b,
            VBValue::Empty | VBValue::Null => false,
            VBValue::Number(n) => *n != 0.0,
            VBValue::String(s) => !s.is_empty(),
            VBValue::Array(_, _) => true,
            VBValue::Object(_) => true,
        }
    }

    fn negate(val: VBValue) -> VBValue {
        match val {
            VBValue::Number(n) => VBValue::Number(-n),
            VBValue::Empty => VBValue::Number(-0.0),
            VBValue::String(s) => {
                let n: f64 = s.parse().unwrap_or(0.0);
                VBValue::Number(-n)
            }
            _ => VBValue::Number(-0.0),
        }
    }

    fn logical_not(val: VBValue) -> VBValue {
        VBValue::Boolean(!Vm::is_truthy(&val))
    }

    fn add(l: VBValue, r: VBValue) -> VBValue {
        if matches!(&l, VBValue::String(_)) || matches!(&r, VBValue::String(_)) {
            Vm::concat_str(l, r)
        } else {
            let ln = value_utils::to_arg_f64(&l);
            let rn = value_utils::to_arg_f64(&r);
            VBValue::Number(ln + rn)
        }
    }

    fn sub(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| a - b)
    }

    fn mul(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| a * b)
    }

    fn div(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| a / b)
    }

    fn int_div(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| (a / b).floor())
    }

    fn mod_op(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| a % b)
    }

    fn pow_op(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_binop(l, r, |a, b| a.powf(b))
    }

    fn number_binop(l: VBValue, r: VBValue, f: fn(f64, f64) -> f64) -> VBValue {
        let ln = value_utils::to_arg_f64(&l);
        let rn = value_utils::to_arg_f64(&r);
        VBValue::Number(f(ln, rn))
    }

    fn concat_str(l: VBValue, r: VBValue) -> VBValue {
        let ls = value_utils::to_arg_string(&l);
        let rs = value_utils::to_arg_string(&r);
        VBValue::String(format!("{}{}", ls, rs).into())
    }

    fn values_equal(a: &VBValue, b: &VBValue) -> bool {
        match (a, b) {
            (VBValue::String(a), VBValue::String(b)) => a == b,
            (VBValue::Number(a), VBValue::Number(b)) => (a - b).abs() < f64::EPSILON,
            (VBValue::Boolean(a), VBValue::Boolean(b)) => a == b,
            (VBValue::Null, VBValue::Null) => true,
            (VBValue::Empty, VBValue::Empty) => true,
            (VBValue::Array(a, _), VBValue::Array(b, _)) => a == b,
            (VBValue::Object(_), VBValue::Object(_)) => false,
            _ => {
                let sa = value_utils::to_arg_string(a);
                let sb = value_utils::to_arg_string(b);
                sa == sb
            }
        }
    }

    fn compare_lt(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_comparison(l, r, |a, b| a < b)
    }

    fn compare_le(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_comparison(l, r, |a, b| a <= b)
    }

    fn compare_gt(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_comparison(l, r, |a, b| a > b)
    }

    fn compare_ge(l: VBValue, r: VBValue) -> VBValue {
        Vm::number_comparison(l, r, |a, b| a >= b)
    }

    fn number_comparison(l: VBValue, r: VBValue, f: fn(f64, f64) -> bool) -> VBValue {
        let ln = value_utils::to_arg_f64(&l);
        let rn = value_utils::to_arg_f64(&r);
        VBValue::Boolean(f(ln, rn))
    }

    fn bool_or_bitwise(l: VBValue, r: VBValue, f: fn(i64, i64) -> i64) -> VBValue {
        match (&l, &r) {
            (VBValue::Boolean(a), VBValue::Boolean(b)) => VBValue::Boolean(f(*a as i64, *b as i64) != 0),
            _ => {
                let ln = value_utils::to_arg_f64(&l) as i64;
                let rn = value_utils::to_arg_f64(&r) as i64;
                VBValue::Number(f(ln, rn) as f64)
            }
        }
    }

    fn imp_op(l: VBValue, r: VBValue) -> VBValue {
        Vm::bool_or_bitwise(l, r, |a, b| !a | b)
    }

    fn eqv_op(l: VBValue, r: VBValue) -> VBValue {
        Vm::bool_or_bitwise(l, r, |a, b| !(a ^ b))
    }

    fn like_match(l: VBValue, r: VBValue) -> VBValue {
        let s = value_utils::to_arg_string(&l);
        let pattern = value_utils::to_arg_string(&r);
        VBValue::Boolean(Vm::like_match_str(&s, &pattern))
    }

    fn like_match_str(s: &str, pattern: &str) -> bool {
        let s_chars: Vec<char> = s.chars().collect();
        let p_chars: Vec<char> = pattern.chars().collect();
        Vm::like_match_rec(&s_chars, &p_chars, 0, 0)
    }

    fn like_match_rec(s: &[char], p: &[char], si: usize, pi: usize) -> bool {
        if pi >= p.len() {
            return si >= s.len();
        }
        match p[pi] {
            '*' => {
                if pi + 1 >= p.len() {
                    return true;
                }
                let mut i = si;
                while i <= s.len() {
                    if Vm::like_match_rec(s, p, i, pi + 1) {
                        return true;
                    }
                    i += 1;
                }
                false
            }
            '?' => {
                if si < s.len() {
                    Vm::like_match_rec(s, p, si + 1, pi + 1)
                } else {
                    false
                }
            }
            '#' => {
                if si < s.len() && s[si].is_ascii_digit() {
                    Vm::like_match_rec(s, p, si + 1, pi + 1)
                } else {
                    false
                }
            }
            '[' => {
                let close = p[pi + 1..].iter().position(|&c| c == ']');
                if let Some(end) = close {
                    let end = end + pi + 1;
                    let negate = pi + 1 < p.len() && p[pi + 1] == '!';
                    let start = if negate { pi + 2 } else { pi + 1 };
                    if si < s.len() {
                        let char_matches = p[start..end].iter().any(|&c| c == s[si]);
                        if char_matches != negate {
                            Vm::like_match_rec(s, p, si + 1, end + 1)
                        } else {
                            false
                        }
                    } else {
                        false
                    }
                } else {
                    si < s.len() && s[si] == '[' && Vm::like_match_rec(s, p, si + 1, pi + 1)
                }
            }
            c => {
                if si < s.len() && (s[si] == c || c.eq_ignore_ascii_case(&s[si])) {
                    Vm::like_match_rec(s, p, si + 1, pi + 1)
                } else {
                    false
                }
            }
        }
    }

    fn is_past_end(counter: &VBValue, end: &VBValue, step: &VBValue) -> bool {
        let c = value_utils::to_arg_f64(counter);
        let e = value_utils::to_arg_f64(end);
        let s = value_utils::to_arg_f64(step);
        if s >= 0.0 {
            c > e
        } else {
            c < e
        }
    }
}
