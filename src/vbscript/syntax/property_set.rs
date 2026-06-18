use super::super::expr::{evaluate, Expr};
use super::super::value::VBValue;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::super::ExecutionContext;
use super::VBSyntax;

#[derive(Clone)]
pub struct PropertySet {
    object_name: String,
    property: String,
    value_expr: Expr,
}

impl PropertySet {
    pub fn new(object_name: String, property: String, value_expr: Expr) -> Self {
        PropertySet {
            object_name,
            property,
            value_expr,
        }
    }
}

impl VBSyntax for PropertySet {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let value = evaluate(&self.value_expr, context)?;

        if self.object_name == "__with_obj__" {
            // With-block property set: swap with_object out, modify, put back
            let mut obj_val = context.scope.with_object.take().ok_or_else(|| {
                VBSErrorType::RuntimeError.into_error("With object not set".to_string())
            })?;
            let result = match &mut obj_val {
                VBValue::Object(ref mut obj) => obj.set_property(&self.property, value, context),
                _ => {
                    return Err(VBSErrorType::RuntimeError
                        .into_error("With object is not an object".to_string()));
                }
            };
            context.scope.with_object = Some(obj_val);
            return result;
        }

        let obj_key = self.object_name.to_uppercase();

        // Take the object out of context, replacing with Empty temporarily
        let mut obj_val = {
            let slot: &mut VBValue = match context.get_variable_mut(&self.object_name) {
                Some(v @ VBValue::Object(_)) => v,
                _ => {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("Object variable '{}' is not set", self.object_name)));
                }
            };
            let mut replacement = VBValue::Empty;
            std::mem::swap(slot, &mut replacement);
            replacement
        };

        // Now the borrow on slot is dropped. obj_val owns the VBValue::Object.
        let result = match &mut obj_val {
            VBValue::Object(ref mut obj) => {
                obj.set_property(&self.property, value, context)
            }
            _ => unreachable!(),
        };

        // Put the object back (even on error, to avoid corrupting state)
        context.set_variable(&obj_key, obj_val);
        result
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
