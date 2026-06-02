use ahash::AHashMap;
use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::value_utils;
use super::vbs_error::{VBSError, VBSErrorType};

#[allow(dead_code)]
/// Trait for VBScript COM / intrinsic objects that can expose properties,
/// methods, and indexed access to scripts.
pub trait VBScriptObject: std::fmt::Debug + Send + Sync {
    /// Clone the object into a new boxed trait object.
    fn clone_box(&self) -> Box<dyn VBScriptObject>;
    /// Return a human-readable type name for debugging.
    fn type_name(&self) -> &'static str { "VBScriptObject" }
    /// Get a named property value.
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError>;
    /// Set a named property value.
    fn set_property(&mut self, _name: &str, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Object does not support setting properties".to_string()
        ))
    }
    fn call_method(&mut self, name: &str, _args: &[VBValue], _context: &mut ExecutionContext) -> Result<VBValue, VBSError>;
    fn indexed_get(&self, _index: &VBValue, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Object does not support indexed access".to_string()
        ))
    }
    fn indexed_set(&mut self, _index: &VBValue, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            "Object does not support indexed access".to_string()
        ))
    }
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

    fn call_method(&mut self, name: &str, args: &[VBValue], _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "ADD" => {
                if args.len() < 2 {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Add requires 2 arguments (key, value)".to_string()
                    ));
                }
                let key = value_utils::to_arg_string(&args[0]);
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
                let key = value_utils::to_arg_string(&args[0]);
                self.items.remove(&key);
                Ok(VBValue::Empty)
            }
            "EXISTS" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Exists requires 1 argument (key)".to_string()
                    ));
                }
                let key = value_utils::to_arg_string(&args[0]);
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

    fn indexed_get(&self, index: &VBValue, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        self.items.get(&key).cloned().ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error(
                format!("Key '{}' not found in Dictionary", key)
            )
        })
    }

    fn indexed_set(&mut self, index: &VBValue, value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        let key = value_utils::to_arg_string(index);
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

    fn call_method(&mut self, name: &str, _args: &[VBValue], _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Method '{}' not found on class '{}'", name, self.class_name)
        ))
    }

    fn indexed_get(&self, _index: &VBValue, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Class '{}' does not support indexed access", self.class_name)
        ))
    }

    fn indexed_set(&mut self, _index: &VBValue, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
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
            "NUMBER" => Ok(VBValue::Number(context.scope.err_number)),
            "DESCRIPTION" => Ok(VBValue::String(context.scope.err_description.clone())),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Property '{}' not found on Err object", name)
            )),
        }
    }

    fn call_method(&mut self, name: &str, args: &[VBValue], _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CLEAR" => Ok(VBValue::Empty),
            "RAISE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Err.Raise requires at least 1 argument (number)".to_string()
                    ));
                }
                let number = value_utils::to_arg_f64(&args[0]) as i32;
                let description = if args.len() > 1 {
                    value_utils::to_arg_string(&args[1])
                } else {
                    "".to_string()
                };
                Err(VBSErrorType::RuntimeError.into_error(description).with_code(number))
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Method '{}' not found on Err object", name)
            )),
        }
    }
}


