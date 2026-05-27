use ahash::AHashMap;
use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::vbs_error::{VBSError, VBSErrorType};

#[allow(dead_code)]
pub trait VBScriptObject: std::fmt::Debug + Send + Sync {
    fn clone_box(&self) -> Box<dyn VBScriptObject>;
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError>;
    fn set_property(&mut self, name: &str, value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError>;
    fn call_method(&mut self, name: &str, args: &[VBValue]) -> Result<VBValue, VBSError>;
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError>;
    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError>;
}

// ---- Dictionary ----

#[derive(Debug, Clone)]
pub struct Dictionary {
    items: AHashMap<String, VBValue>,
}

impl Dictionary {
    pub fn new() -> Self {
        Dictionary { items: AHashMap::new() }
    }
}

impl VBScriptObject for Dictionary {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.items.len() as f64)),
            "KEYS" => Ok(VBValue::Array(std::sync::Arc::new(
                self.items.keys().map(|k| VBValue::String(k.clone())).collect()
            ))),
            "ITEMS" => Ok(VBValue::Array(std::sync::Arc::new(
                self.items.values().cloned().collect()
            ))),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Property '{}' not found on Dictionary", name)
            )),
        }
    }

    fn set_property(&mut self, name: &str, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Cannot set property '{}' on Dictionary", name)
            )),
        }
    }

    fn call_method(&mut self, name: &str, args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "ADD" => {
                if args.len() < 2 {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Add requires 2 arguments (key, value)".to_string()
                    ));
                }
                let key = to_arg_string(&args[0]);
                let value = args[1].clone();
                self.items.insert(key, value);
                Ok(VBValue::Empty)
            }
            "REMOVE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Remove requires 1 argument (key)".to_string()
                    ));
                }
                let key = to_arg_string(&args[0]);
                self.items.remove(&key);
                Ok(VBValue::Empty)
            }
            "EXISTS" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Exists requires 1 argument (key)".to_string()
                    ));
                }
                let key = to_arg_string(&args[0]);
                Ok(VBValue::Boolean(self.items.contains_key(&key)))
            }
            "REMOVEALL" => {
                self.items.clear();
                Ok(VBValue::Empty)
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Method '{}' not found on Dictionary", name)
            )),
        }
    }

    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = to_arg_string(index);
        self.items.get(&key).cloned().ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error(
                format!("Key '{}' not found in Dictionary", key)
            )
        })
    }

    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError> {
        let key = to_arg_string(index);
        self.items.insert(key, value);
        Ok(())
    }
}

// ---- ClassInstance ----

#[derive(Debug)]
pub struct ClassInstance {
    pub class_name: String,
    pub instance_vars: AHashMap<String, VBValue>,
}

impl ClassInstance {
    pub fn new(class_name: &str) -> Self {
        ClassInstance {
            class_name: class_name.to_string(),
            instance_vars: AHashMap::new(),
        }
    }
}

impl VBScriptObject for ClassInstance {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(ClassInstance {
            class_name: self.class_name.clone(),
            instance_vars: self.instance_vars.clone(),
        })
    }

    fn get_property(&self, name: &str, context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        let class_def = context.get_class(&self.class_name)
            .ok_or_else(|| VBSErrorType::RuntimeError.into_error(
                format!("Class '{}' not found", self.class_name)
            ))?;
        let prop_name_upper = name.to_uppercase();
        if let Some(prop_def) = class_def.properties.get(&prop_name_upper) {
            if let Some(ref body_lines) = prop_def.get_body {
                let body_blocks = super::block::parse_blocks(body_lines)
                    .map_err(|_| VBSErrorType::RuntimeError.into_error(
                        format!("Error parsing Property Get '{}' body", name)
                    ))?;
                let mut instance_vars = self.instance_vars.clone();
                context.set_variable(name, VBValue::Empty);
                instance_vars.insert(name.to_uppercase(), VBValue::Empty);
                let result = context.with_instance_scope(&mut instance_vars, |ctx| {
                    super::block::execute_blocks(&body_blocks, ctx)
                });
                match result {
                    Ok(()) => {
                        let val = instance_vars.get(&prop_name_upper)
                            .cloned()
                            .unwrap_or(VBValue::Empty);
                        Ok(val)
                    }
                    Err(e) => Err(e),
                }
            } else {
                let val = self.instance_vars.get(&prop_name_upper)
                    .cloned()
                    .unwrap_or(VBValue::Empty);
                Ok(val)
            }
        } else {
            let val = self.instance_vars.get(&prop_name_upper)
                .cloned()
                .unwrap_or(VBValue::Empty);
            Ok(val)
        }
    }

    fn set_property(&mut self, name: &str, value: VBValue, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let class_def = context.get_class(&self.class_name)
            .ok_or_else(|| VBSErrorType::RuntimeError.into_error(
                format!("Class '{}' not found", self.class_name)
            ))?;
        let prop_name_upper = name.to_uppercase();
        if let Some(prop_def) = class_def.properties.get(&prop_name_upper) {
            if let Some(ref body_lines) = prop_def.let_body {
                let body_blocks = super::block::parse_blocks(body_lines)
                    .map_err(|_| VBSErrorType::RuntimeError.into_error(
                        format!("Error parsing Property Let '{}' body", name)
                    ))?;
                let mut instance_vars = std::mem::take(&mut self.instance_vars);
                if let Some(ref param) = prop_def.let_param {
                    instance_vars.insert(param.to_uppercase(), value.clone());
                }
                let result = context.with_instance_scope(&mut instance_vars, |ctx| {
                    super::block::execute_blocks(&body_blocks, ctx)
                });
                self.instance_vars = instance_vars;
                result
            } else {
                self.instance_vars.insert(prop_name_upper, value);
                Ok(())
            }
        } else {
            self.instance_vars.insert(prop_name_upper, value);
            Ok(())
        }
    }

    fn call_method(&mut self, name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Method '{}' not found on class '{}'", name, self.class_name)
        ))
    }

    fn indexed_get(&self, _index: &VBValue) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Class '{}' does not support indexed access", self.class_name)
        ))
    }

    fn indexed_set(&mut self, _index: &VBValue, _value: VBValue) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Class '{}' does not support indexed access", self.class_name)
        ))
    }
}

// ---- ErrObject ----

#[derive(Debug, Clone)]
pub struct ErrObject;

impl ErrObject {
    pub fn new() -> Self {
        ErrObject
    }
}

impl VBScriptObject for ErrObject {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "NUMBER" => Ok(VBValue::Number(context.err_number)),
            "DESCRIPTION" => Ok(VBValue::String(context.err_description.clone())),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Property '{}' not found on Err object", name)
            )),
        }
    }

    fn set_property(&mut self, _name: &str, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Cannot set properties on Err object".to_string()
        ))
    }

    fn call_method(&mut self, name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CLEAR" => Ok(VBValue::Empty),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Method '{}' not found on Err object", name)
            )),
        }
    }

    fn indexed_get(&self, _index: &VBValue) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Err object does not support indexed access".to_string()
        ))
    }

    fn indexed_set(&mut self, _index: &VBValue, _value: VBValue) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Err object does not support indexed access".to_string()
        ))
    }
}

fn to_arg_string(val: &VBValue) -> String {
    match val {
        VBValue::String(s) => s.clone(),
        VBValue::Null => "Null".to_string(),
        VBValue::Empty => "".to_string(),
        VBValue::Number(n) => n.to_string(),
        VBValue::Boolean(true) => "True".to_string(),
        VBValue::Boolean(false) => "False".to_string(),
        VBValue::Array(_) => "Array".to_string(),
        VBValue::Object(_) => "Object".to_string(),
    }
}
