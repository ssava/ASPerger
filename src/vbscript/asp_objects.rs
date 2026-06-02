use ahash::AHashMap;
use std::sync::Mutex;

use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::value_utils;
use super::vbobject::VBScriptObject;
use super::vbs_error::{VBSError, VBSErrorType};

// ===== Global Stores =====

static SESSION_STORE: std::sync::OnceLock<Mutex<AHashMap<String, AHashMap<String, VBValue>>>> =
    std::sync::OnceLock::new();

pub fn get_session_store() -> &'static Mutex<AHashMap<String, AHashMap<String, VBValue>>> {
    SESSION_STORE.get_or_init(|| Mutex::new(AHashMap::new()))
}

static APPLICATION_STORE: std::sync::OnceLock<Mutex<AHashMap<String, VBValue>>> =
    std::sync::OnceLock::new();

pub fn get_app_store() -> &'static Mutex<AHashMap<String, VBValue>> {
    APPLICATION_STORE.get_or_init(|| Mutex::new(AHashMap::new()))
}

#[allow(dead_code)]
pub fn clear_app_store() {
    if let Some(store) = APPLICATION_STORE.get() {
        store.lock().unwrap().clear();
    }
}

static APP_LOCK: std::sync::OnceLock<std::sync::Mutex<()>> = std::sync::OnceLock::new();

fn get_app_lock() -> &'static std::sync::Mutex<()> {
    APP_LOCK.get_or_init(|| std::sync::Mutex::new(()))
}

// ===== RequestObject =====

#[derive(Debug, Clone)]
pub struct RequestObject;

impl VBScriptObject for RequestObject {
    fn type_name(&self) -> &'static str {
        "Request"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "QUERYSTRING" => Ok(VBValue::Object(Box::new(RequestQueryString(
                context.request_params.clone(),
            )))),
            "FORM" => Ok(VBValue::Object(Box::new(RequestForm(
                context.request_form.clone(),
            )))),
            "SERVERVARIABLES" => Ok(VBValue::Object(Box::new(RequestServerVariables(
                context.request_headers.clone(),
            )))),
            "COOKIES" => Ok(VBValue::Object(Box::new(RequestCookies(
                context.request_cookies.clone(),
            )))),
            "TOTALBYTES" => Ok(VBValue::Number(context.request_total_bytes as f64)),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Request",
                name
            ))),
        }
    }

    fn call_method(&mut self, name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "BINARYREAD" => Ok(VBValue::Empty),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Method '{}' not found on Request",
                name
            ))),
        }
    }
}

// ===== Request Sub-Collections =====

#[derive(Debug, Clone)]
pub struct RequestQueryString(pub AHashMap<String, String>);

impl VBScriptObject for RequestQueryString {
    fn type_name(&self) -> &'static str {
        "RequestQueryString"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on RequestQueryString",
                name
            ))),
        }
    }
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val))
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub struct RequestForm(pub AHashMap<String, String>);

impl VBScriptObject for RequestForm {
    fn type_name(&self) -> &'static str {
        "RequestForm"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on RequestForm",
                name
            ))),
        }
    }
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val))
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub struct RequestServerVariables(pub AHashMap<String, String>);

impl VBScriptObject for RequestServerVariables {
    fn type_name(&self) -> &'static str {
        "RequestServerVariables"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => {
                let val = self.0.get(&name.to_lowercase()).cloned().unwrap_or_default();
                Ok(VBValue::String(val))
            }
        }
    }
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key.to_lowercase()).cloned().unwrap_or_default();
        Ok(VBValue::String(val))
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

#[derive(Debug, Clone)]
pub struct RequestCookies(pub AHashMap<String, String>);

impl VBScriptObject for RequestCookies {
    fn type_name(&self) -> &'static str {
        "RequestCookies"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => Ok(VBValue::Number(self.0.len() as f64)),
            _ => {
                let val = self.0.get(name).cloned().unwrap_or_default();
                Ok(VBValue::String(val))
            }
        }
    }
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let val = self.0.get(&key).cloned().unwrap_or_default();
        Ok(VBValue::String(val))
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

// ===== ResponseObject =====

#[derive(Debug, Clone)]
pub struct ResponseObject;

impl VBScriptObject for ResponseObject {
    fn type_name(&self) -> &'static str {
        "Response"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "BUFFER" => Ok(VBValue::Boolean(true)),
            "CONTENTTYPE" => Ok(VBValue::String("text/html".to_string())),
            "STATUS" => Ok(VBValue::String(context.response_status.clone())),
            "EXPIRES" => Ok(VBValue::Number(0.0)),
            "COOKIES" => Ok(VBValue::Object(Box::new(ResponseCookies::new()))),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Response",
                name
            ))),
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
                context
                    .response_extra_headers
                    .push(("Content-Type".to_string(), value_utils::to_arg_string(&value)));
                Ok(())
            }
            "STATUS" => {
                context.response_status = value_utils::to_arg_string(&value);
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Response",
                name
            ))),
        }
    }

    fn call_method(&mut self, name: &str, args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "WRITE" => {
                if !args.is_empty() {
                    // Note: The syntax shortcut handles most Response.Write calls,
                    // but this supports the method call style as well.
                    // The shortcut writes directly to response_buffer before we get here.
                }
                Ok(VBValue::Empty)
            }
            "REDIRECT" => {
                Ok(VBValue::Empty)
            }
            "END" => Ok(VBValue::Empty),
            "CLEAR" => Ok(VBValue::Empty),
            "FLUSH" => Ok(VBValue::Empty),
            "ADDHEADER" => Ok(VBValue::Empty),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Method '{}' not found on Response",
                name
            ))),
        }
    }
}

// ===== ResponseCookies =====

#[derive(Debug, Clone)]
pub struct ResponseCookies {
    cookies: AHashMap<String, String>,
}

impl ResponseCookies {
    pub fn new() -> Self {
        ResponseCookies {
            cookies: AHashMap::new(),
        }
    }
}

impl VBScriptObject for ResponseCookies {
    fn type_name(&self) -> &'static str {
        "ResponseCookies"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        let val = self.cookies.get(name).cloned().unwrap_or_default();
        Ok(VBValue::String(val))
    }
    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError> {
        let name = value_utils::to_arg_string(index);
        let val = value_utils::to_arg_string(&value);
        self.cookies.insert(name, val);
        Ok(())
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

// ===== SessionObject =====

#[derive(Debug, Clone)]
pub struct SessionObject {
    pub session_id: String,
    pub session_enabled: bool,
}

impl VBScriptObject for SessionObject {
    fn type_name(&self) -> &'static str {
        "Session"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        match name.to_uppercase().as_str() {
            "SESSIONID" => Ok(VBValue::String(context.session_id.clone())),
            "TIMEOUT" => Ok(VBValue::Number(20.0)),
            "CONTENTS" => Ok(VBValue::Object(Box::new(SessionContents::new(
                context.session_id.clone(),
            )))),
            _ => {
                let store = get_session_store().lock().unwrap();
                if let Some(data) = store.get(&context.session_id.to_uppercase()) {
                    if let Some(val) = data.get(&name.to_uppercase()) {
                        return Ok(val.clone());
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
                let mut store = get_session_store().lock().unwrap();
                store
                    .entry(context.session_id.to_uppercase())
                    .or_default()
                    .insert(name.to_uppercase(), value);
                Ok(())
            }
        }
    }

    fn call_method(&mut self, name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        match name.to_uppercase().as_str() {
            "ABANDON" => {
                let mut store = get_session_store().lock().unwrap();
                store.remove(&self.session_id.to_uppercase());
                Ok(VBValue::Empty)
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Method '{}' not found on Session",
                name
            ))),
        }
    }

    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        if !self.session_enabled {
            return Ok(VBValue::Empty);
        }
        let key = value_utils::to_arg_string(index);
        let store = get_session_store().lock().unwrap();
        if let Some(data) = store.get(&self.session_id.to_uppercase()) {
            Ok(data
                .get(&key.to_uppercase())
                .cloned()
                .unwrap_or(VBValue::Empty))
        } else {
            Ok(VBValue::Empty)
        }
    }

    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError> {
        if !self.session_enabled {
            return Ok(());
        }
        let key = value_utils::to_arg_string(index);
        let mut store = get_session_store().lock().unwrap();
        store
            .entry(self.session_id.to_uppercase())
            .or_default()
            .insert(key.to_uppercase(), value);
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct SessionContents {
    session_id: String,
}

impl SessionContents {
    pub fn new(session_id: String) -> Self {
        SessionContents { session_id }
    }
}

impl VBScriptObject for SessionContents {
    fn type_name(&self) -> &'static str {
        "SessionContents"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => {
                let store = get_session_store().lock().unwrap();
                if let Some(data) = store.get(&self.session_id.to_uppercase()) {
                    Ok(VBValue::Number(data.len() as f64))
                } else {
                    Ok(VBValue::Number(0.0))
                }
            }
            _ => {
                let store = get_session_store().lock().unwrap();
                if let Some(data) = store.get(&self.session_id.to_uppercase()) {
                    Ok(data
                        .get(&name.to_uppercase())
                        .cloned()
                        .unwrap_or(VBValue::Empty))
                } else {
                    Ok(VBValue::Empty)
                }
            }
        }
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
}

// ===== ServerObject =====

#[derive(Debug, Clone)]
pub struct ServerObject;

impl VBScriptObject for ServerObject {
    fn type_name(&self) -> &'static str {
        "Server"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "SCRIPTPATH" => Ok(VBValue::String("".to_string())),
            "SCRIPTTIMEOUT" => Ok(VBValue::Number(90.0)),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Server",
                name
            ))),
        }
    }

    fn call_method(&mut self, name: &str, args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CREATEOBJECT" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Server.CreateObject requires 1 argument".to_string(),
                    ));
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
                    "ADODB.CONNECTION" => Ok(VBValue::Object(Box::new(
                        super::adodb::Connection::new(),
                    ))),
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
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Method '{}' not found on Server",
                name
            ))),
        }
    }
}

// ===== ApplicationObject =====

#[derive(Debug, Clone)]
pub struct ApplicationObject;

impl VBScriptObject for ApplicationObject {
    fn type_name(&self) -> &'static str {
        "Application"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CONTENTS" => Ok(VBValue::Object(Box::new(ApplicationContents))),
            "STATICOBJECTS" => Ok(VBValue::Empty),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Application",
                name
            ))),
        }
    }

    fn call_method(&mut self, name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "LOCK" => {
                let _guard = get_app_lock().lock().unwrap();
                Ok(VBValue::Empty)
            }
            "UNLOCK" => Ok(VBValue::Empty),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Method '{}' not found on Application",
                name
            ))),
        }
    }

    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let store = get_app_store().lock().unwrap();
        Ok(store
            .get(&key.to_uppercase())
            .cloned()
            .unwrap_or(VBValue::Empty))
    }

    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError> {
        let key = value_utils::to_arg_string(index);
        let mut store = get_app_store().lock().unwrap();
        store.insert(key.to_uppercase(), value);
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct ApplicationContents;

impl VBScriptObject for ApplicationContents {
    fn type_name(&self) -> &'static str {
        "ApplicationContents"
    }
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }
    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "COUNT" => {
                let store = get_app_store().lock().unwrap();
                Ok(VBValue::Number(store.len() as f64))
            }
            _ => {
                let store = get_app_store().lock().unwrap();
                Ok(store
                    .get(&name.to_uppercase())
                    .cloned()
                    .unwrap_or(VBValue::Empty))
            }
        }
    }
    fn call_method(&mut self, _name: &str, _args: &[VBValue]) -> Result<VBValue, VBSError> {
        Ok(VBValue::Empty)
    }
    fn indexed_get(&self, index: &VBValue) -> Result<VBValue, VBSError> {
        let key = value_utils::to_arg_string(index);
        let store = get_app_store().lock().unwrap();
        Ok(store
            .get(&key.to_uppercase())
            .cloned()
            .unwrap_or(VBValue::Empty))
    }
    fn indexed_set(&mut self, index: &VBValue, value: VBValue) -> Result<(), VBSError> {
        let key = value_utils::to_arg_string(index);
        let mut store = get_app_store().lock().unwrap();
        store.insert(key.to_uppercase(), value);
        Ok(())
    }
}
