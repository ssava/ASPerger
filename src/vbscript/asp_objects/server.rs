use super::super::execution_context::ExecutionContext;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbobject::VBScriptObject;
use super::super::vbs_error::{VBSError, VBSErrorType};
use crate::{impl_vbscript_object, prop_not_found, method_not_found, cannot_set_property};

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
                        super::super::vbobject::Dictionary::new(),
                    ))),
                    "SCRIPTING.FILESYSTEMOBJECT" => Ok(VBValue::Object(Box::new(
                        super::super::fso::FileSystemObject::new(),
                    ))),
                    "VBSCRIPT.REGEXP" => Ok(VBValue::Object(Box::new(
                        super::super::regexp::RegExpObject::new(),
                    ))),
                    "ADODB.CONNECTION" => {
                        Ok(VBValue::Object(Box::new(super::super::adodb::Connection::new())))
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
