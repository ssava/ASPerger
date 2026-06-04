use super::super::expr::{evaluate, Expr};
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::super::ExecutionContext;
use super::VBSyntax;

pub struct MethodCall {
    object_name: String,
    method_name: String,
    args: Vec<Expr>,
}

impl MethodCall {
    pub fn new(object_name: String, method_name: String, args: Vec<Expr>) -> Self {
        MethodCall {
            object_name,
            method_name,
            args,
        }
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
        context.scope.with_object.clone().ok_or_else(|| {
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
                        if let Ok(result) = sub_obj.indexed_get(&args[0], context) {
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
        let args: Result<Vec<VBValue>, VBSError> =
            self.args.iter().map(|arg| evaluate(arg, context)).collect();
        let args = args?;

        // Handle Response methods that need context access
        if self.object_name.eq_ignore_ascii_case("response") {
            match self.method_name.to_uppercase().as_str() {
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
                    return Ok(());
                }
                "END" => {
                    context.response.ended = true;
                    return Ok(());
                }
                "CLEAR" => {
                    context.response.buffer.clear();
                    return Ok(());
                }
                "FLUSH" => {
                    context.response.flushed.push_str(&context.response.buffer);
                    context.response.buffer.clear();
                    return Ok(());
                }
                "ADDHEADER" => {
                    if args.len() >= 2 {
                        let name = value_utils::to_arg_string(&args[0]);
                        let value = value_utils::to_arg_string(&args[1]);
                        context.response.extra_headers.push((name, value));
                    }
                    return Ok(());
                }
                _ => {}
            }
        }

        // Handle Server methods
        if self.object_name.eq_ignore_ascii_case("server") {
            match self.method_name.to_uppercase().as_str() {
                "EXECUTE" | "TRANSFER" => {
                    if !args.is_empty() {
                        let path = value_utils::to_arg_string(&args[0]);
                        let callback = context.execute_file_callback.take();
                        if let Some(cb) = callback {
                            cb(&path, context).map_err(|e| {
                                VBSErrorType::RuntimeError
                                    .into_error(format!("Server.Execute failed: {e}"))
                            })?;
                            context.execute_file_callback = Some(cb);
                        }
                        if self.method_name.to_uppercase().as_str() == "TRANSFER" {
                            context.response.ended = true;
                        }
                    }
                    return Ok(());
                }
                _ => {}
            }
        }

        // ASP pattern: obj.Property(args) — try property + indexed_get first
        if !args.is_empty() && self.object_name != "__with_obj__" {
            if try_property_indexed_access(&self.object_name, &self.method_name, &args, context)?
                .is_some()
            {
                return Ok(());
            }
        }

        if self.object_name == "__with_obj__" {
            // With-block method call: use context.scope.with_object
            let mut obj_val = context.scope.with_object.take().ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("With object not set".to_string())
            })?;
            match &mut obj_val {
                VBValue::Object(ref mut obj) => {
                    obj.call_method(&self.method_name, &args, context)?;
                }
                _ => {
                    context.scope.with_object = Some(obj_val);
                    return Err(VBSErrorType::RuntimeError
                        .into_error("With object is not an object".to_string()));
                }
            }
            context.scope.with_object = Some(obj_val);
            return Ok(());
        }

        match context.get_variable_mut(&self.object_name) {
            Some(v @ VBValue::Object(_)) => {
                let mut obj_val = VBValue::Empty;
                std::mem::swap(v, &mut obj_val);
                let result = match &mut obj_val {
                    VBValue::Object(ref mut obj) => {
                        obj.call_method(&self.method_name, &args, context)
                    }
                    _ => unreachable!(),
                };
                context.set_variable(&self.object_name.to_uppercase(), obj_val);
                result.map(|_| ())
            }
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Object variable '{}' is not set", self.object_name))),
        }
    }
}
