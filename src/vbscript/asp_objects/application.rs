use super::super::execution_context::ExecutionContext;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbobject::VBScriptObject;
use super::super::vbs_error::VBSError;
use crate::{impl_vbscript_object, prop_not_found, method_not_found};

#[derive(Debug, Clone)]
pub struct ApplicationObject;

impl VBScriptObject for ApplicationObject {
    impl_vbscript_object!(ApplicationObject, "Application");

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CONTENTS" => Ok(VBValue::Object(Box::new(ApplicationContents))),
            "STATICOBJECTS" => Ok(VBValue::Empty),
            _ => prop_not_found!("Application", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        _args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "LOCK" => {
                if let Some(ref store) = context.store {
                    store.lock_app_blocking(context.request_id);
                }
                Ok(VBValue::Empty)
            }
            "UNLOCK" => {
                if let Some(ref store) = context.store {
                    store.unlock_app(context.request_id);
                }
                Ok(VBValue::Empty)
            }
            _ => method_not_found!("Application", name),
        }
    }

    fn indexed_get(
        &self,
        index: &VBValue,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            store.wait_for_app_unlock(context.request_id);
            let apps = store.lock_apps();
            return Ok(apps
                .get(&key.to_uppercase())
                .cloned()
                .unwrap_or(VBValue::Empty));
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
            store.wait_for_app_unlock(context.request_id);
            let mut apps = store.lock_apps();
            apps.insert(key.to_uppercase(), value);
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ApplicationContents;

impl VBScriptObject for ApplicationContents {
    impl_vbscript_object!(ApplicationContents, "ApplicationContents");
    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        if let Some(ref store) = context.store {
            store.wait_for_app_unlock(context.request_id);
            let apps = store.lock_apps();
            match name.to_uppercase().as_str() {
                "COUNT" => Ok(VBValue::Number(apps.len() as f64)),
                _ => Ok(apps
                    .get(&name.to_uppercase())
                    .cloned()
                    .unwrap_or(VBValue::Empty)),
            }
        } else {
            Ok(VBValue::Number(0.0))
        }
    }
    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        if let Some(ref store) = context.store {
            store.wait_for_app_unlock(context.request_id);
            let apps = store.lock_apps();
            return Ok(apps
                .get(&key.to_uppercase())
                .cloned()
                .unwrap_or(VBValue::Empty));
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
            store.wait_for_app_unlock(context.request_id);
            let mut apps = store.lock_apps();
            apps.insert(key.to_uppercase(), value);
        }
        Ok(())
    }
}
