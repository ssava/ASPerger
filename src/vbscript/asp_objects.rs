//! ASP intrinsic object implementations: `Request`, `Response`, `Session`,
//! `Server`, `Application`, and their sub-collections (`QueryString`, `Form`,
//! `ServerVariables`, `Cookies`, `Contents`, `StaticObjects`).
//!
//! These objects are injected into `ExecutionContext` as global variables
//! before any script runs (see `AspServer::inject_asp_intrinsic_objects`).

use ahash::AHashMap;

use super::execution_context::{CookieEntry, ExecutionContext};
use super::value::VBValue;
use super::value_utils;
use super::vbobject::VBScriptObject;
use super::vbs_error::{VBSError, VBSErrorType};
use crate::{impl_vbscript_object, prop_not_found, method_not_found, cannot_set_property};

// ===== RequestObject =====

/// ASP intrinsic `Request` object.
///
/// Properties: `Request.QueryString`, `Request.Form`, `Request.ServerVariables`,
/// `Request.Cookies`, `Request.TotalBytes`.
/// Methods: `Request.BinaryRead`.
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

// ===== Request Sub-Collections =====

#[derive(Debug, Clone)]
/// `Request.QueryString` — read-only indexed collection of URL query parameters.
pub struct RequestQueryString(pub AHashMap<String, String>);

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
        Ok(VBValue::String(val))
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
/// `Request.Form` — read-only indexed collection of POST form fields.
pub struct RequestForm(pub AHashMap<String, String>);

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
        Ok(VBValue::String(val))
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
/// `Request.ServerVariables` — read-only indexed collection of
/// CGI-style server variables (HTTP headers, connection info).
pub struct RequestServerVariables(pub AHashMap<String, String>);

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
                Ok(VBValue::String(val))
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
        Ok(VBValue::String(val))
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
/// `Request.Cookies` — read-only indexed collection of cookie values.
pub struct RequestCookies(pub AHashMap<String, String>);

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
                Ok(VBValue::String(val))
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
        Ok(VBValue::String(val))
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

// ===== ResponseObject =====

/// ASP intrinsic `Response` object.
///
/// Methods: `Response.Write`, `Response.End`, `Response.Clear`,
/// `Response.Flush`, `Response.Redirect`, `Response.AddHeader`,
/// `Response.BinaryWrite`.
/// Properties: `Response.Buffer`, `Response.ContentType`, `Response.Status`,
/// `Response.Cookies`, `Response.IsClientConnected`.
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
            "CONTENTTYPE" => Ok(VBValue::String("text/html".to_string())),
            "STATUS" => Ok(VBValue::String(context.response.status.clone())),
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
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "WRITE" => {
                if !args.is_empty() {
                    // Note: The syntax shortcut handles most Response.Write calls,
                    // but this supports the method call style as well.
                    // The shortcut writes directly to response_buffer before we get here.
                }
                Ok(VBValue::Empty)
            }
            "REDIRECT" => Ok(VBValue::Empty),
            "END" => Ok(VBValue::Empty),
            "CLEAR" => Ok(VBValue::Empty),
            "FLUSH" => Ok(VBValue::Empty),
            "ADDHEADER" => Ok(VBValue::Empty),
            _ => method_not_found!("Response", name),
        }
    }
}

// ===== ResponseCookies =====

/// `Response.Cookies` — collection proxy for `Response.Cookies("name")` access.
/// Writes go through `ResponseCookiesSet` / `ResponseCookiesSetProp` syntax nodes.
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

/// A single cookie with its attributes and subkeys.
/// Stateless — reads/writes through `context.response.cookies[cookie_name]`.
#[derive(Debug, Clone)]
/// `Response.Cookies(key)` — handles indexed get/set on the cookies
/// collection, returning a `CookieEntry` wrapper for sub-property access
/// like `Response.Cookies("name").Expires = ...`.
pub struct CookieObject {
    cookie_name: String,
}

impl CookieObject {
    pub fn new(cookie_name: String) -> Self {
        CookieObject { cookie_name }
    }
}

/// Serialise a cookie name+entry into a `Set-Cookie` header value.
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
            "EXPIRES" => Ok(VBValue::String(entry.expires.clone())),
            "DOMAIN" => Ok(VBValue::String(entry.domain.clone())),
            "PATH" => Ok(VBValue::String(entry.path.clone())),
            "SECURE" => Ok(VBValue::Boolean(entry.secure)),
            "HASKEYS" => Ok(VBValue::Boolean(!entry.subkeys.is_empty())),
            _ => Ok(VBValue::String(entry.value.to_string())),
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
                .map(VBValue::String)
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
                Some(entry) => Ok(VBValue::String(entry.value.to_string())),
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

// ===== SessionObject =====

/// ASP intrinsic `Session` object.
///
/// Properties: `Session.SessionID`, `Session.Timeout`.
/// Indexed access: `Session("key")` — reads/writes session variables
/// via the shared `Store`.
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
            "SESSIONID" => Ok(VBValue::String(context.session.id.clone())),
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
/// `Session.Contents` — read-only indexed collection of session variables.
pub struct SessionContents {
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
                        Ok(VBValue::String(keys[index - 1].clone()))
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

// ===== ServerObject =====

/// ASP intrinsic `Server` object.
///
/// Methods: `Server.HTMLEncode`, `Server.URLEncode`, `Server.Execute`,
/// `Server.Transfer`, `Server.CreateObject`.
/// Properties: `Server.ScriptTimeout`.
#[derive(Debug, Clone)]
pub struct ServerObject;

impl VBScriptObject for ServerObject {
    impl_vbscript_object!(ServerObject, "Server");

    fn get_property(
        &self,
        name: &str,
        context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "SCRIPTPATH" => Ok(VBValue::String(context.script_path.clone())),
            "SCRIPTTIMEOUT" => Ok(VBValue::Number(90.0)),
            _ => prop_not_found!("Server", name),
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        _value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            "SCRIPTTIMEOUT" => Ok(()),
            _ => cannot_set_property!("Server", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CREATEOBJECT" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError
                        .into_error("Server.CreateObject requires 1 argument".to_string()));
                }
                let prog_id = value_utils::to_arg_string(&args[0]);
                match prog_id.to_uppercase().as_str() {
                    "SCRIPTING.DICTIONARY" => Ok(VBValue::Object(Box::new(
                        super::vbobject::Dictionary::new(),
                    ))),
                    "SCRIPTING.FILESYSTEMOBJECT" => Ok(VBValue::Object(Box::new(
                        super::fso::FileSystemObject::new(),
                    ))),
                    "VBSCRIPT.REGEXP" => Ok(VBValue::Object(Box::new(
                        super::regexp::RegExpObject::new(),
                    ))),
                    "ADODB.CONNECTION" => {
                        Ok(VBValue::Object(Box::new(super::adodb::Connection::new())))
                    }
                    _ => Err(VBSErrorType::NotImplementedError.into_error(format!(
                        "Server.CreateObject('{}') is not implemented",
                        prog_id
                    ))),
                }
            }
            "MAPPATH" => {
                let path = value_utils::to_arg_string(&args[0]);
                let cwd = std::env::current_dir().unwrap_or_default();
                let full_path = cwd.join(path.trim_start_matches('/').trim_start_matches('\\'));
                Ok(VBValue::String(
                    full_path.to_str().unwrap_or(&path).to_string(),
                ))
            }
            "HTMLENCODE" => {
                let s = value_utils::to_arg_string(&args[0]);
                let encoded = s
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("\"", "&quot;")
                    .replace("'", "&#39;");
                Ok(VBValue::String(encoded))
            }
            "URLENCODE" => {
                let s = value_utils::to_arg_string(&args[0]);
                let encoded: String = s
                    .bytes()
                    .map(|b| match b {
                        b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => {
                            (b as char).to_string()
                        }
                        b' ' => "+".to_string(),
                        _ => format!("%{:02X}", b),
                    })
                    .collect();
                Ok(VBValue::String(encoded))
            }
            "URLPATHENCODE" => {
                let s = value_utils::to_arg_string(&args[0]);
                let encoded: String = s
                    .bytes()
                    .map(|b| match b {
                        b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' => {
                            (b as char).to_string()
                        }
                        b'/' => "/".to_string(),
                        b' ' => "+".to_string(),
                        _ => format!("%{:02X}", b),
                    })
                    .collect();
                Ok(VBValue::String(encoded))
            }
            _ => method_not_found!("Server", name),
        }
    }
}

// ===== ApplicationObject =====

/// ASP intrinsic `Application` object.
///
/// Methods: `Application.Lock`, `Application.Unlock`.
/// Indexed access: `Application("key")` — reads/writes application-scoped
/// variables via the shared `Store` with `Condvar`-based locking.
/// Sub-collections: `Application.Contents`, `Application.StaticObjects`.
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

    /// Application.Lock / Unlock / Contents
    ///
    /// VBScript concurrency model: `Application.Lock` blocks the calling thread
    /// until the lock is available.  While locked, **other** requests are parked
    /// by `Condvar` when they try to read or write `Application` variables.
    /// The locking request reads/writes freely (reentrant).
    ///
    /// - `lock_app_blocking` — used by `Application.Lock` (writer side)
    /// - `wait_for_app_unlock` — used by every `Application("key")` access
    ///   (reader side — must wait for any outstanding writer lock to clear)
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
            // Wait for any outstanding Application.Lock from another request
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
pub struct ApplicationContents;

impl VBScriptObject for ApplicationContents {
    impl_vbscript_object!(ApplicationContents, "ApplicationContents");
    /// Mirror of `ApplicationObject` locking: every access waits for the
    /// writer lock (`Application.Lock`) held by another request to clear.
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
