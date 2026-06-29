use super::super::execution_context::{CookieEntry, ExecutionContext};
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbobject::VBScriptObject;
use super::super::vbs_error::VBSError;
use crate::{impl_vbscript_object, prop_not_found, method_not_found, cannot_set_property};

#[derive(Debug, Clone)]
pub struct ResponseObject;

impl VBScriptObject for ResponseObject {
    impl_vbscript_object!(ResponseObject, "Response");

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "BUFFER" => Ok(VBValue::Boolean(true)),
            "CONTENTTYPE" => Ok(VBValue::String("text/html".into())),
            "STATUS" => Ok(VBValue::String(context.response.status.clone().into())),
            "EXPIRES" => Ok(VBValue::Number(0.0)),
            "COOKIES" => Ok(VBValue::Object(Box::new(ResponseCookies::new()))),
            _ => prop_not_found!("Response", name),
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            "CONTENTTYPE" => {
                context.response.extra_headers.push((
                    "Content-Type".to_string(),
                    value_utils::to_arg_string(&value),
                ));
                Ok(())
            }
            "STATUS" => {
                context.response.status = value_utils::to_arg_string(&value);
                Ok(())
            }
            "BUFFER" => {
                context.response.buffer = value_utils::to_arg_string(&value);
                Ok(())
            }
            "EXPIRES" => Ok(()),
            _ => cannot_set_property!("Response", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "WRITE" => Ok(VBValue::Empty),
            "REDIRECT" => {
                if !args.is_empty() {
                    let url = value_utils::to_arg_string(&args[0]);
                    context.response.status = "302 Found".to_string();
                    context
                        .response
                        .extra_headers
                        .push(("Location".to_string(), url.clone()));
                    context.response.redirect_url = url;
                    context.response.ended = true;
                }
                Ok(VBValue::Empty)
            }
            "END" => {
                context.response.ended = true;
                Ok(VBValue::Empty)
            }
            "CLEAR" => {
                context.response.buffer.clear();
                Ok(VBValue::Empty)
            }
            "FLUSH" => {
                context.response.flushed.push_str(&context.response.buffer);
                context.response.buffer.clear();
                Ok(VBValue::Empty)
            }
            "ADDHEADER" => {
                if args.len() >= 2 {
                    let name = value_utils::to_arg_string(&args[0]);
                    let value = value_utils::to_arg_string(&args[1]);
                    context.response.extra_headers.push((name, value));
                }
                Ok(VBValue::Empty)
            }
            "BINARYWRITE" => {
                if let Some(arg) = args.first() {
                    let bytes = match arg {
                        VBValue::Array(items, _dims) => {
                            items.iter().map(|v| match v {
                                VBValue::Number(n) => *n as u8,
                                VBValue::Boolean(b) => *b as u8,
                                other => other.to_string().as_bytes().first().copied().unwrap_or(0),
                            }).collect()
                        }
                        VBValue::Null | VBValue::Empty => Vec::new(),
                        other => other.to_string().into_bytes(),
                    };
                    context.response.write_binary(&bytes);
                }
                Ok(VBValue::Empty)
            }
            _ => method_not_found!("Response", name),
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct ResponseCookies;

impl ResponseCookies {
    pub fn new() -> Self {
        ResponseCookies
    }
}

fn get_or_create_entry<'a>(context: &'a mut ExecutionContext, name: &str) -> &'a mut CookieEntry {
    context.response.cookies.entry(name.to_string()).or_default()
}

#[derive(Debug, Clone)]
pub(crate) struct CookieObject {
    cookie_name: String,
}

impl CookieObject {
    pub fn new(cookie_name: String) -> Self {
        CookieObject { cookie_name }
    }
}

pub(crate) fn to_cookie_string(name: &str, entry: &CookieEntry) -> String {
    let mut s = format!("{}={}", name, entry.value);
    if !entry.expires.is_empty() {
        s.push_str(&format!("; expires={}", entry.expires));
    }
    if !entry.domain.is_empty() {
        s.push_str(&format!("; domain={}", entry.domain));
    }
    if !entry.path.is_empty() {
        s.push_str(&format!("; path={}", entry.path));
    } else {
        s.push_str("; path=/");
    }
    if entry.secure {
        s.push_str("; secure");
    }
    s
}

impl VBScriptObject for CookieObject {
    impl_vbscript_object!(CookieObject, "Cookie");
    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let entry = match context.response.cookies.get(&self.cookie_name) {
            Some(e) => e,
            None => return Ok(VBValue::Empty),
        };
        match name.to_uppercase().as_str() {
            "EXPIRES" => Ok(VBValue::String(entry.expires.clone().into())),
            "DOMAIN" => Ok(VBValue::String(entry.domain.clone().into())),
            "PATH" => Ok(VBValue::String(entry.path.clone().into())),
            "SECURE" => Ok(VBValue::Boolean(entry.secure)),
            "HASKEYS" => Ok(VBValue::Boolean(!entry.subkeys.is_empty())),
            _ => Ok(VBValue::String(entry.value.to_string().into())),
        }
    }
    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let entry = get_or_create_entry(context, &self.cookie_name);
        match name.to_uppercase().as_str() {
            "EXPIRES" => entry.expires = value_utils::to_arg_string(&value),
            "DOMAIN" => entry.domain = value_utils::to_arg_string(&value),
            "PATH" => entry.path = value_utils::to_arg_string(&value),
            "SECURE" => entry.secure = value_utils::to_boolean(&value),
            _ => entry.value = value_utils::to_arg_string(&value),
        }
        Ok(())
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        match context.response.cookies.get(&self.cookie_name) {
            Some(entry) => Ok(entry
                .subkeys
                .get(&key.to_uppercase())
                .cloned()
                .map(|s| VBValue::String(s.into()))
                .unwrap_or(VBValue::Empty)),
            None => Ok(VBValue::Empty),
        }
    }
    fn indexed_set(
        &mut self,
        index: &VBValue,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = value_utils::to_arg_string(&value);
        let entry = get_or_create_entry(context, &self.cookie_name);
        entry.subkeys.insert(key.to_uppercase(), val);
        Ok(())
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

impl VBScriptObject for ResponseCookies {
    impl_vbscript_object!(ResponseCookies, "ResponseCookies");
    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(context.response.cookies.len() as f64)),
            _ => match context.response.cookies.get(name) {
                Some(entry) => Ok(VBValue::String(entry.value.to_string().into())),
                None => Ok(VBValue::Empty),
            },
        }
    }
    fn indexed_get(
        &self,
        index: &VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        let name = value_utils::to_arg_string(index);
        Ok(VBValue::Object(Box::new(CookieObject::new(name))))
    }
    fn indexed_set(
        &mut self,
        index: &VBValue,
        value: VBValue,
        context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        let name = value_utils::to_arg_string(index);
        let val = value_utils::to_arg_string(&value);
        let entry = get_or_create_entry(context, &name);
        entry.value = val;
        Ok(())
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
