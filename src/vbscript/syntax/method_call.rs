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
        let args: Result<Vec<VBValue>, VBSError> = self.args.iter()
            .map(|arg| evaluate(arg, context))
            .collect();

        match context.get_variable_mut(&self.object_name) {
            Some(VBValue::Object(ref mut obj)) => {
                obj.call_method(&self.method_name, &args?)?;
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Object variable '{}' is not set", self.object_name)
            )),
        }
    }
}
