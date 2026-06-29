use super::super::execution_context::ExecutionContext;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbobject::VBScriptObject;
use super::super::vbs_error::{VBSError, VBSErrorType};
use crate::{impl_vbscript_object, method_not_found, prop_not_found};

#[derive(Debug, Clone)]
pub struct SessionObject {
    pub session_id: String,
    pub session_enabled: bool,
}

impl VBScriptObject for SessionObject {
    impl_vbscript_object!(SessionObject, "Session");

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        match name.to_uppercase().as_str() {
            "SESSIONID" => Ok(VBValue::String(context.session.id.clone().into())),
            "TIMEOUT" => Ok(VBValue::Number(20.0)),
            "CONTENTS" => Ok(VBValue::Object(Box::new(SessionContents::new(
                context.session.id.clone(),
            )))),
            _ => {
                if let Some(ref store) = context.store {
                    let sessions = store.lock_sessions();
                    if let Some(data) = sessions.get(&context.session.id.to_uppercase()) {
                        if let Some(val) = data.get(&name.to_uppercase()) {
                            return Ok(val.clone());
                        }
                    }
                }
                Ok(VBValue::Empty)
            }
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        if !self.session_enabled {
            return Ok(());
        }
        match name.to_uppercase().as_str() {
            "TIMEOUT" => Ok(()),
            _ => {
                if let Some(ref store) = context.store {
                    let mut sessions = store.lock_sessions();
                    sessions
                        .entry(context.session.id.to_uppercase())
                        .or_default()
                        .insert(name.to_uppercase(), value);
                }
                Ok(())
            }
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        _args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        match name.to_uppercase().as_str() {
            "ABANDON" => {
                if let Some(ref store) = context.store {
                    let mut sessions = store.lock_sessions();
                    sessions.remove(&self.session_id.to_uppercase());
                }
                Ok(VBValue::Empty)
            }
            _ => method_not_found!("Session", name),
        }
    }

    fn indexed_get(
        &self,
        index: &VBValue,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            let sessions = store.lock_sessions();
            if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                return Ok(data
                    .get(&key.to_uppercase())
                    .cloned()
                    .unwrap_or(VBValue::Empty));
            }
        }
        Ok(VBValue::Empty)
    }

    fn indexed_set(
        &mut self,
        index: &VBValue,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        if !self.session_enabled {
            return Ok(());
        }
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            let mut sessions = store.lock_sessions();
            sessions
                .entry(self.session_id.to_uppercase())
                .or_default()
                .insert(key.to_uppercase(), value);
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct SessionContents {
    session_id: String,
}

impl SessionContents {
    pub fn new(session_id: String) -> Self {
        SessionContents { session_id }
    }
}

impl VBScriptObject for SessionContents {
    impl_vbscript_object!(SessionContents, "SessionContents");
    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => {
                if let Some(ref store) = context.store {
                    let sessions = store.lock_sessions();
                    if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                        return Ok(VBValue::Number(data.len() as f64));
                    }
                }
                Ok(VBValue::Number(0.0))
            }
            "KEY" | "ITEM" | "REMOVE" | "REMOVEALL" => prop_not_found!("SessionContents", name),
            _ => {
                if let Some(ref store) = context.store {
                    let sessions = store.lock_sessions();
                    if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                        return Ok(data
                            .get(&name.to_uppercase())
                            .cloned()
                            .unwrap_or(VBValue::Empty));
                    }
                }
                Ok(VBValue::Empty)
            }
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            let sessions = store.lock_sessions();
            if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                return Ok(data
                    .get(&key.to_uppercase())
                    .cloned()
                    .unwrap_or(VBValue::Empty));
            }
        }
        Ok(VBValue::Empty)
    }
    fn indexed_set(
        &mut self,
        index: &VBValue,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            let mut sessions = store.lock_sessions();
            sessions
                .entry(self.session_id.to_uppercase())
                .or_default()
                .insert(key.to_uppercase(), value);
        }
        Ok(())
    }
    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        if let Some(ref store) = context.store {
            match name.to_uppercase().as_str() {
                "KEY" => {
                    let sessions = store.lock_sessions();
                    if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                        if args.is_empty() {
                            return Err(VBSErrorType::RuntimeError.into_error(
                                "Session.Contents.Key requires 1 argument (index)".to_string(),
                            ));
                        }
                        let index = value_utils::to_arg_f64(&args[0]) as usize;
                        let keys: Vec<&String> = data.keys().collect();
                        if index < 1 || index > keys.len() {
                            return Err(VBSErrorType::RuntimeError.into_error(format!(
                                "Key index out of range: {} (valid: 1-{})",
                                index,
                                keys.len()
                            )));
                        }
                        Ok(VBValue::String(keys[index - 1].clone().into()))
                    } else {
                        Ok(VBValue::Empty)
                    }
                }
                "ITEM" => {
                    let sessions = store.lock_sessions();
                    if let Some(data) = sessions.get(&self.session_id.to_uppercase()) {
                        if args.is_empty() {
                            return Err(VBSErrorType::RuntimeError.into_error(
                                "Session.Contents.Item requires 1 argument (index)".to_string(),
                            ));
                        }
                        let index = value_utils::to_arg_f64(&args[0]) as usize;
                        let values: Vec<VBValue> = data.values().cloned().collect();
                        if index < 1 || index > values.len() {
                            return Err(VBSErrorType::RuntimeError.into_error(format!(
                                "Item index out of range: {} (valid: 1-{})",
                                index,
                                values.len()
                            )));
                        }
                        Ok(values[index - 1].clone())
                    } else {
                        Ok(VBValue::Empty)
                    }
                }
                "REMOVE" => {
                    if args.is_empty() {
                        return Err(VBSErrorType::RuntimeError.into_error(
                            "Session.Contents.Remove requires 1 argument (key)".to_string(),
                        ));
                    }
                    let key = value_utils::to_arg_string(&args[0]);
                    let mut sessions = store.lock_sessions();
                    if let Some(data) = sessions.get_mut(&self.session_id.to_uppercase()) {
                        data.remove(&key.to_uppercase());
                    }
                    Ok(VBValue::Empty)
                }
                "REMOVEALL" => {
                    let mut sessions = store.lock_sessions();
                    sessions.remove(&self.session_id.to_uppercase());
                    Ok(VBValue::Empty)
                }
                _ => Ok(VBValue::Empty),
            }
        } else {
            Ok(VBValue::Empty)
        }
    }
}
