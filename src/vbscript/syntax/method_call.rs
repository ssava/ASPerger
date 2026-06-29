use super::super::compiler::Compiler;
use super::super::expr::{evaluate, Expr};
use super::super::instruction::Instruction;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::super::ExecutionContext;
use super::VBSyntax;

/// AST node for `obj.Method(args)` statements.
///
/// Dispatches to one of several call paths depending on the object:
/// - `Response.End`, `Response.Redirect`, etc. — short-circuits with side effects
/// - `Server.Execute` / `Server.Transfer` — calls the `execute_file_callback`
/// - `obj.Property(args)` — tries property access + indexed_get pattern first
/// - `With obj ... .Method(args)` — uses `context.with_object`
/// - Generic `obj.Method(args)` — swaps the object out of context, calls `call_method`, puts it back
#[derive(Clone)]
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

/// Try the ASP intrinsic pattern `obj.Property(args)`.
///
/// Some ASP objects expose collections through a property that itself
/// supports indexed access — e.g. `Request.QueryString("id")` or
/// `Session.Contents("key")`.  This helper peels off one level of
/// property access then applies `indexed_get` with the first argument.
///
/// Returns `Some(value)` on success, `None` if the pattern doesn't match.
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
    let obj_ref = if object_name == "__with_obj__" {
        context
            .with_object
            .take()
            .ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("With object not set".to_string())
            })?
    } else {
        let slot = context
            .get_variable_mut(object_name)
            .and_then(|v| match v {
                VBValue::Object(_) => Some(v),
                _ => None,
            });
        match slot {
            Some(slot) => {
                let mut replacement = VBValue::Empty;
                std::mem::swap(slot, &mut replacement);
                replacement
            }
            None => return Ok(None),
        }
    };
    if let VBValue::Object(ref obj) = obj_ref {
        if let Ok(VBValue::Object(sub_obj)) = obj.get_property(property, context).as_ref() {
            if let Ok(result) = sub_obj.indexed_get(&args[0], context) {
                if object_name == "__with_obj__" {
                    context.with_object = Some(obj_ref);
                } else {
                    context.set_variable(object_name, obj_ref);
                }
                return Ok(Some(result));
            }
        }
    }
    if object_name == "__with_obj__" {
        context.with_object = Some(obj_ref);
    } else {
        context.set_variable(object_name, obj_ref);
    }
    Ok(None)
}

fn write_binary_value(value: &VBValue) -> Vec<u8> {
    match value {
        VBValue::Array(items, _dims) => {
            items.iter().map(|v| match v {
                VBValue::Number(n) => *n as u8,
                VBValue::Boolean(b) => *b as u8,
                other => other.to_string().as_bytes().first().copied().unwrap_or(0),
            }).collect()
        }
        VBValue::Null | VBValue::Empty => Vec::new(),
        other => other.to_string().into_bytes(),
    }
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
                "BINARYWRITE" => {
                    if let Some(arg) = args.first() {
                        let bytes = write_binary_value(arg);
                        context.response.write_binary(&bytes);
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
        if !args.is_empty() && self.object_name != "__with_obj__"
            && try_property_indexed_access(&self.object_name, &self.method_name, &args, context)?
                .is_some()
            {
                return Ok(());
            }

        if self.object_name == "__with_obj__" {
            // With-block method call: use context.with_object
            let mut obj_val = context.with_object.take().ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("With object not set".to_string())
            })?;
            match &mut obj_val {
                VBValue::Object(ref mut obj) => {
                    obj.call_method(&self.method_name, &args, context)?;
                }
                _ => {
                    context.with_object = Some(obj_val);
                    return Err(VBSErrorType::RuntimeError
                        .into_error("With object is not an object".to_string()));
                }
            }
            context.with_object = Some(obj_val);
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

    fn compile(&self, compiler: &mut Compiler) -> Result<(), VBSError> {
        let name_lower = self.object_name.to_lowercase();
        let method_idx = compiler.add_constant(VBValue::String(self.method_name.to_lowercase().into()));
        for arg in &self.args {
            compiler.compile_expr(arg);
        }
        if let Some(slot) = compiler.local_slot(&name_lower) {
            compiler.emit(Instruction::CallMethodLocal(slot, method_idx, self.args.len() as u8));
        } else {
            let obj_idx = compiler.add_constant(VBValue::String(name_lower.into()));
            compiler.emit(Instruction::CallMethodGlobal(obj_idx, method_idx, self.args.len() as u8));
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
