use super::VBSyntax;
use super::super::expr::{evaluate, Expr};
use super::super::value::VBValue;
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

impl VBSyntax for MethodCall {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let mut obj_val = context.get_variable(&self.object_name).ok_or_else(|| {
            VBSErrorType::RuntimeError.into_error(
                format!("Object variable '{}' is not set", self.object_name)
            )
        })?;

        match &mut obj_val {
            VBValue::Object(ref mut obj) => {
                let args: Result<Vec<VBValue>, VBSError> = self.args.iter()
                    .map(|arg| evaluate(arg, context))
                    .collect();
                obj.call_method(&self.method_name, &args?)?;
                context.set_variable(&self.object_name, obj_val);
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                "Object doesn't support this property or method".to_string()
            )),
        }
    }
}
