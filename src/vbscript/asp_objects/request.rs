use ahash::AHashMap;

use super::super::execution_context::ExecutionContext;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbobject::VBScriptObject;
use super::super::vbs_error::VBSError;
use crate::{impl_vbscript_object, prop_not_found, method_not_found};

#[derive(Debug, Clone)]
pub struct RequestObject;

impl VBScriptObject for RequestObject {
    impl_vbscript_object!(RequestObject, "Request");

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "QUERYSTRING" => Ok(VBValue::Object(Box::new(RequestQueryString(
                context.request.params.clone(),
            )))),
            "FORM" => Ok(VBValue::Object(Box::new(RequestForm(
                context.request.form.clone(),
            )))),
            "SERVERVARIABLES" => Ok(VBValue::Object(Box::new(RequestServerVariables(
                context.request.headers.clone(),
            )))),
            "COOKIES" => Ok(VBValue::Object(Box::new(RequestCookies(
                context.request.cookies.clone(),
            )))),
            "TOTALBYTES" => Ok(VBValue::Number(context.request.total_bytes as f64)),
            _ => prop_not_found!("Request", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "BINARYREAD" => Ok(VBValue::Empty),
            _ => method_not_found!("Request", name),
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RequestQueryString(pub AHashMap<String, String>);

impl VBScriptObject for RequestQueryString {
    impl_vbscript_object!(RequestQueryString, "RequestQueryString");
    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => prop_not_found!("RequestQueryString", name),
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val.into()))
    }
    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RequestForm(pub AHashMap<String, String>);

impl VBScriptObject for RequestForm {
    impl_vbscript_object!(RequestForm, "RequestForm");
    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => prop_not_found!("RequestForm", name),
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val.into()))
    }
    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RequestServerVariables(pub AHashMap<String, String>);

impl VBScriptObject for RequestServerVariables {
    impl_vbscript_object!(RequestServerVariables, "RequestServerVariables");
    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => {
                let val = self
                    .0
                    .get(&name.to_lowercase())
                    .cloned()
                    .unwrap_or_default();
                Ok(VBValue::String(val.into()))
            }
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key.to_lowercase()).cloned().unwrap_or_default();
        Ok(VBValue::String(val.into()))
    }
    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RequestCookies(pub AHashMap<String, String>);

impl VBScriptObject for RequestCookies {
    impl_vbscript_object!(RequestCookies, "RequestCookies");
    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => {
                let val = self.0.get(name).cloned().unwrap_or_default();
                Ok(VBValue::String(val.into()))
            }
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val.into()))
    }
    fn call_method(
        &mut self,
        _name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}
