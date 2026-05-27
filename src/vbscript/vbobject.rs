use ahash::AHashMap;
use super::value::VBValue;
use super::vbs_error::{VBSError, VBSErrorType};

#[allow(dead_code)]
pub trait VBScriptObject: std::fmt::Debug + Send + Sync {
    fn clone_box(&self) -> Box<dyn VBScriptObject>;
    fn get_property(&self, name: &str) -> Result<VBValue, VBSError>;
    fn set_property(&mut self, name: &str, value: VBValue) -> Result<(), VBSError>;
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

    fn get_property(&self, name: &str) -> Result<VBValue, VBSError> {
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

    fn set_property(&mut self, name: &str, _value: VBValue) -> Result<(), VBSError> {
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
