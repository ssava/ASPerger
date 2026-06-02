use super::VBSyntax;
use super::super::expr::{evaluate, Expr};
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::super::ExecutionContext;

pub struct MethodCall {
    object_name: String,
    method_name: String,
    args: Vec<Expr>,
}

impl MethodCall {
    pub fn new(object_name: String, method_name: String, args: Vec<Expr>) -> Self {
        MethodCall { object_name, method_name, args }
    }
}

fn try_property_indexed_access(
    object_name: &str,
    property: &str,
    args: &[VBValue],
    context: &mut ExecutionContext,
) -> Result<Option<VBValue>, VBSError> {
    // Check if this is a property access + indexed_get pattern
    // e.g., Request.QueryString "name" or Session.Contents("key")
    if args.is_empty() {
        return Ok(None);
    }
    let obj_val = if object_name == "__with_obj__" {
        context.with_object.clone().ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error("With object not set".to_string())
        })?
    } else {
        match context.get_variable(object_name) {
            Some(VBValue::Object(_)) => context.get_variable(object_name).unwrap().clone(),
            _ => return Ok(None),
        }
    };
    match &obj_val {
        VBValue::Object(obj) => {
            if let Ok(prop_val) = obj.get_property(property, context) {
                match &prop_val {
                    VBValue::Object(sub_obj) => {
                        if let Ok(result) = sub_obj.indexed_get(&args[0]) {
                            return Ok(Some(result));
                        }
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
    Ok(None)
}

impl VBSyntax for MethodCall {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let args: Result<Vec<VBValue>, VBSError> = self.args.iter()
            .map(|arg| evaluate(arg, context))
            .collect();
        let args = args?;

        // Handle Response methods that need context access
        if self.object_name.eq_ignore_ascii_case("response") {
            match self.method_name.to_uppercase().as_str() {
                "REDIRECT" => {
                    if !args.is_empty() {
                        let url = value_utils::to_arg_string(&args[0]);
                        context.response_status = "302 Found".to_string();
                        context.response_extra_headers
                            .push(("Location".to_string(), url.clone()));
                        context.response_redirect_url = url;
                        context.response_ended = true;
                    }
                    return Ok(());
                }
                "END" => {
                    context.response_ended = true;
                    return Ok(());
                }
                "CLEAR" => {
                    context.response_buffer.clear();
                    return Ok(());
                }
                "FLUSH" => {
                    context.response_flushed.push_str(&context.response_buffer);
                    context.response_buffer.clear();
                    return Ok(());
                }
                "ADDHEADER" => {
                    if args.len() >= 2 {
                        let name = value_utils::to_arg_string(&args[0]);
                        let value = value_utils::to_arg_string(&args[1]);
                        context.response_extra_headers.push((name, value));
                    }
                    return Ok(());
                }
                _ => {}
            }
        }

        // ASP pattern: obj.Property(args) — try property + indexed_get first
        if !args.is_empty() && self.object_name != "__with_obj__" {
            if try_property_indexed_access(
                &self.object_name,
                &self.method_name,
                &args,
                context,
            )?.is_some() {
                return Ok(());
            }
        }

        if self.object_name == "__with_obj__" {
            // With-block method call: use context.with_object
            match context.with_object.as_mut() {
                Some(VBValue::Object(ref mut obj)) => {
                    obj.call_method(&self.method_name, &args)?;
                    return Ok(());
                }
                Some(_) => {
                    return Err(VBSErrorType::RuntimeError.into_error(
                        "With object is not an object".to_string()
                    ));
                }
                None => {
                    return Err(VBSErrorType::RuntimeError.into_error(
                        "With object not set".to_string()
                    ));
                }
            }
        }

        match context.get_variable_mut(&self.object_name) {
            Some(VBValue::Object(ref mut obj)) => {
                obj.call_method(&self.method_name, &args)?;
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Object variable '{}' is not set", self.object_name)
            )),
        }
    }
}
