use std::borrow::Cow;

use super::block::{parse_blocks, execute_blocks};
use super::execution_context::{CIStr, CIString, ExecutionContext};
use super::value::VBValue;
use super::value_utils;
use super::vbs_error::{VBSError, VBSErrorType};
use ahash::AHashMap;

#[allow(dead_code)]
/// Trait for VBScript COM / intrinsic objects that can expose properties,
/// methods, and indexed access to scripts.
pub trait VBScriptObject: std::fmt::Debug + Send + Sync {
    /// Clone the object into a new boxed trait object.
    fn clone_box(&self) -> Box<dyn VBScriptObject>;
    /// Return a human-readable type name for debugging.
    fn type_name(&self) -> &'static str {
        "VBScriptObject"
    }
    /// Get a named property value.
    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError>;
    /// Set a named property value.
    fn set_property(
        &mut self,
        _name: &str,
        _value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError
            .into_error("Object does not support setting properties".to_string()))
    }
    fn call_method(
        &mut self,
        name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError>;
    fn indexed_get(
        &self,
        _index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError
            .into_error("Object does not support indexed access".to_string()))
    }
    fn indexed_set(
        &mut self,
        _index: &VBValue,
        _value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError
            .into_error("Object does not support indexed access".to_string()))
    }
}

// ---- Dictionary ----

#[derive(Debug, Clone)]
pub struct Dictionary {
    items: AHashMap<String, VBValue>,
}

impl Default for Dictionary {
    fn default() -> Self {
        Self::new()
    }
}

impl Dictionary {
    pub fn new() -> Self {
        Dictionary {
            items: AHashMap::new(),
        }
    }
}

/// Like `to_arg_string` but returns `Cow<str>` to avoid allocation
/// when the value is already a `String`.
fn key_to_cow(val: &VBValue) -> Cow<'_, str> {
    match val {
        VBValue::String(s) => Cow::Borrowed(s.as_str()),
        VBValue::Null => Cow::Owned("Null".to_string()),
        VBValue::Empty => Cow::Owned(String::new()),
        VBValue::Number(n) => Cow::Owned(n.to_string()),
        VBValue::Boolean(true) => Cow::Owned("True".to_string()),
        VBValue::Boolean(false) => Cow::Owned("False".to_string()),
        VBValue::Array(..) => Cow::Owned("Array".to_string()),
        VBValue::Object(_) => Cow::Owned("Object".to_string()),
    }
}

impl VBScriptObject for Dictionary {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.items.len() as f64)),
            "KEYS" => Ok(VBValue::Array(std::sync::Arc::new(
                self.items
                    .keys()
                    .map(|k| VBValue::String(k.clone()))
                    .collect(),
            ), vec![])),
            "ITEMS" => Ok(VBValue::Array(std::sync::Arc::new(
                self.items.values().cloned().collect(),
            ), vec![])),
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Property '{}' not found on Dictionary", name))),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "ADD" => {
                if args.len() < 2 {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Dictionary.Add requires 2 arguments (key, value)".to_string(),
                    ));
                }
                let key = key_to_cow(&args[0]).into_owned();
                let value = args[1].clone();
                self.items.insert(key, value);
                Ok(VBValue::Empty)
            }
            "REMOVE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("Dictionary.Remove requires 1 argument (key)".to_string()));
                }
                let key = key_to_cow(&args[0]);
                self.items.remove(key.as_ref());
                Ok(VBValue::Empty)
            }
            "EXISTS" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("Dictionary.Exists requires 1 argument (key)".to_string()));
                }
                let key = key_to_cow(&args[0]);
                Ok(VBValue::Boolean(self.items.contains_key(key.as_ref())))
            }
            "ITEM" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("Dictionary.Item requires 1 argument (key)".to_string()));
                }
                let key = key_to_cow(&args[0]);
                self.items.get(key.as_ref()).cloned().ok_or_else(|| {
                    VBSErrorType::RuntimeError
                        .into_error(format!("Key '{}' not found in Dictionary", key))
                })
            }
            "REMOVEALL" => {
                self.items.clear();
                Ok(VBValue::Empty)
            }
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Method '{}' not found on Dictionary", name))),
        }
    }

    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = key_to_cow(index);
        self.items.get(key.as_ref()).cloned().ok_or_else(|| {
            VBSErrorType::RuntimeError
                .into_error(format!("Key '{}' not found in Dictionary", key))
        })
    }

    fn indexed_set(
        &mut self,
        index: &VBValue,
        value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let key = key_to_cow(index).into_owned();
        self.items.insert(key, value);
        Ok(())
    }
}

#[cfg(test)]
mod dictionary_tests {
    use super::*;
    use crate::vbscript::execution_context::ExecutionContext;

    fn ctx() -> ExecutionContext {
        ExecutionContext::new()
    }

    #[test]
    fn test_dictionary_new_empty() {
        let d = Dictionary::new();
        assert_eq!(d.items.len(), 0);
    }

    #[test]
    fn test_dictionary_add_and_indexed_get() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("key1".into()), VBValue::String("val1".into())], &mut c).unwrap();
        let val = d.indexed_get(&VBValue::String("key1".into()), &mut c).unwrap();
        assert_eq!(val, VBValue::String("val1".into()));
    }

    #[test]
    fn test_dictionary_add_duplicate_key_overwrites() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("k".into()), VBValue::Number(1.0)], &mut c).unwrap();
        d.call_method("ADD", &[VBValue::String("k".into()), VBValue::Number(2.0)], &mut c).unwrap();
        let val = d.indexed_get(&VBValue::String("k".into()), &mut c).unwrap();
        assert_eq!(val, VBValue::Number(2.0));
    }

    #[test]
    fn test_dictionary_count() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        assert_eq!(d.get_property("COUNT", &mut c).unwrap(), VBValue::Number(0.0));
        d.call_method("ADD", &[VBValue::String("a".into()), VBValue::Empty], &mut c).unwrap();
        assert_eq!(d.get_property("COUNT", &mut c).unwrap(), VBValue::Number(1.0));
    }

    #[test]
    fn test_dictionary_exists() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("k".into()), VBValue::Empty], &mut c).unwrap();
        let exists = d.call_method("EXISTS", &[VBValue::String("k".into())], &mut c).unwrap();
        assert_eq!(exists, VBValue::Boolean(true));
        let not_exists = d.call_method("EXISTS", &[VBValue::String("missing".into())], &mut c).unwrap();
        assert_eq!(not_exists, VBValue::Boolean(false));
    }

    #[test]
    fn test_dictionary_keys() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("x".into()), VBValue::Number(1.0)], &mut c).unwrap();
        let keys = d.get_property("KEYS", &mut c).unwrap();
        match keys {
            VBValue::Array(ref v, _) => {
                assert_eq!(v.len(), 1);
                assert_eq!(v[0], VBValue::String("x".into()));
            }
            _ => panic!("expected Array"),
        }
    }

    #[test]
    fn test_dictionary_remove_all() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("a".into()), VBValue::Empty], &mut c).unwrap();
        d.call_method("ADD", &[VBValue::String("b".into()), VBValue::Empty], &mut c).unwrap();
        d.call_method("REMOVEALL", &[], &mut c).unwrap();
        assert_eq!(d.get_property("COUNT", &mut c).unwrap(), VBValue::Number(0.0));
    }

    #[test]
    fn test_dictionary_remove() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        d.call_method("ADD", &[VBValue::String("k".into()), VBValue::Number(1.0)], &mut c).unwrap();
        d.call_method("REMOVE", &[VBValue::String("k".into())], &mut c).unwrap();
        assert_eq!(d.get_property("COUNT", &mut c).unwrap(), VBValue::Number(0.0));
    }

    #[test]
    fn test_dictionary_add_requires_two_args() {
        let mut d = Dictionary::new();
        let mut c = ctx();
        let result = d.call_method("ADD", &[VBValue::String("k".into())], &mut c);
        assert!(result.is_err());
    }
}

// ---- ClassInstance ----

#[derive(Debug)]
pub struct ClassInstance {
    pub class_name: String,
    pub instance_vars: AHashMap<CIString, VBValue>,
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

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let class_def = context.get_class(&self.class_name).ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error(format!("Class '{}' not found", self.class_name))
        })?;
        if let Some(prop_def) = class_def.properties.get(name) {
            if let Some(ref body_lines) = prop_def.get_body {
                let body_blocks = super::block::parse_blocks(body_lines).map_err(|_| {
                    VBSErrorType::RuntimeError
                        .into_error(format!("Error parsing Property Get '{}' body", name))
                })?;
                let mut instance_vars = self.instance_vars.clone();
                context.set_variable(name, VBValue::Empty);
                instance_vars.insert(CIString::new(name.to_string()), VBValue::Empty);
                let result = context.with_instance_scope(&mut instance_vars, |ctx| {
                    super::block::execute_blocks(&body_blocks, ctx)
                });
                match result {
                    Ok(()) => {
                        let val = instance_vars
                            .get(CIStr::new(name))
                            .cloned()
                            .unwrap_or(VBValue::Empty);
                        Ok(val)
                    }
                    Err(e) => Err(e),
                }
            } else {
                let val = self
                    .instance_vars
                    .get(CIStr::new(name))
                    .cloned()
                    .unwrap_or(VBValue::Empty);
                Ok(val)
            }
        } else {
            let val = self
                .instance_vars
                .get(CIStr::new(name))
                .cloned()
                .unwrap_or(VBValue::Empty);
            Ok(val)
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let class_def = context.get_class(&self.class_name).ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error(format!("Class '{}' not found", self.class_name))
        })?;
        if let Some(prop_def) = class_def.properties.get(name) {
            if let Some(ref body_lines) = prop_def.let_body {
                let body_blocks = super::block::parse_blocks(body_lines).map_err(|_| {
                    VBSErrorType::RuntimeError
                        .into_error(format!("Error parsing Property Let '{}' body", name))
                })?;
                let mut instance_vars = std::mem::take(&mut self.instance_vars);
                if let Some(ref param) = prop_def.let_param {
                    instance_vars.insert(CIString::new(param.clone()), value.clone());
                }
                let result = context.with_instance_scope(&mut instance_vars, |ctx| {
                    super::block::execute_blocks(&body_blocks, ctx)
                });
                self.instance_vars = instance_vars;
                result
            } else {
                self.instance_vars.insert(CIString::new(name.to_string()), value);
                Ok(())
            }
        } else {
            self.instance_vars.insert(CIString::new(name.to_string()), value);
            Ok(())
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        // Clone method data upfront to avoid borrow conflicts with context
        let method = context
            .get_class(&self.class_name)
            .and_then(|c| {
                // Try exact match first, then case-insensitive lookup
                c.methods.get(name).cloned().or_else(|| {
                    let upper = name.to_uppercase();
                    c.methods.get(&upper).cloned()
                })
            })
            .ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error(format!(
                    "Method '{}' not found on class '{}'",
                    name, self.class_name
                ))
            })?;

        // Parse and cache method body
        let cache_key = format!("__cls_{}_{}", self.class_name, method.name);
        let body_blocks = match context.get_function_body(&cache_key) {
            Some(cached) => cached.clone(),
            None => {
                let blocks = parse_blocks(&method.body_lines).map_err(|_| {
                    VBSErrorType::RuntimeError.into_error(format!(
                        "Error parsing method '{}' body",
                        method.name
                    ))
                })?;
                let name = cache_key.clone();
                context.set_function_body(&name, blocks.clone());
                blocks
            }
        };

        // Build instance vars map with method params
        let mut instance_vars = self.instance_vars.clone();
        for (i, param) in method.params.iter().enumerate() {
            let val = args.get(i).cloned().unwrap_or(VBValue::Empty);
            instance_vars.insert(CIString::new(param.clone()), val);
        }

        // For Functions, init return value variable
        if method.is_function {
            instance_vars.insert(CIString::new(method.name.clone()), VBValue::Empty);
        }

        // Execute with merged scope (globals + instance vars)
        let result = context.with_class_method_scope(&mut instance_vars, |ctx| {
            execute_blocks(&body_blocks, ctx)
        });

        // Capture updated instance vars
        self.instance_vars = instance_vars;

        match result {
            Ok(()) => {
                if method.is_function {
                    Ok(self
                        .instance_vars
                        .get(CIStr::new(&method.name))
                        .cloned()
                        .unwrap_or(VBValue::Empty))
                } else {
                    Ok(VBValue::Empty)
                }
            }
            Err(e) if e.is_exit_function() || e.is_exit_sub() => {
                if method.is_function {
                    Ok(self
                        .instance_vars
                        .get(CIStr::new(&method.name))
                        .cloned()
                        .unwrap_or(VBValue::Empty))
                } else {
                    Ok(VBValue::Empty)
                }
            }
            Err(e) => Err(e),
        }
    }

    fn indexed_get(
        &self,
        _index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(format!(
            "Class '{}' does not support indexed access",
            self.class_name
        )))
    }

    fn indexed_set(
        &mut self,
        _index: &VBValue,
        _value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(format!(
            "Class '{}' does not support indexed access",
            self.class_name
        )))
    }
}

// ---- ErrObject ----

#[derive(Debug, Clone)]
pub struct ErrObject;

impl Default for ErrObject {
    fn default() -> Self {
        Self::new()
    }
}

impl ErrObject {
    pub fn new() -> Self {
        ErrObject
    }
}

impl VBScriptObject for ErrObject {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "NUMBER" => Ok(VBValue::Number(context.scope.err_number)),
            "DESCRIPTION" => Ok(VBValue::String(context.scope.err_description.clone())),
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Property '{}' not found on Err object", name))),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CLEAR" => {
                context.clear_err();
                Ok(VBValue::Empty)
            }
            "RAISE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Err.Raise requires at least 1 argument (number)".to_string(),
                    ));
                }
                let number = value_utils::to_arg_f64(&args[0]) as i32;
                let description = if args.len() > 1 {
                    value_utils::to_arg_string(&args[1])
                } else {
                    "".to_string()
                };
                Err(VBSErrorType::RuntimeError
                    .into_error(description)
                    .with_code(number))
            }
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Method '{}' not found on Err object", name))),
        }
    }
}
