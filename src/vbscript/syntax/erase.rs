use super::VBSyntax;
use crate::vbscript::value::VBValue;
use crate::vbscript::vbs_error::VBSError;
use crate::vbscript::ExecutionContext;

#[derive(Clone)]
pub struct Erase {
    var_names: Vec<String>,
}

impl Erase {
    pub fn new(var_names: Vec<String>) -> Self {
        Erase { var_names }
    }
}

impl VBSyntax for Erase {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        for name in &self.var_names {
            if let Some(v) = context.get_variable_mut(name) {
                match v {
                    VBValue::Array(ref mut items, _) => {
                        let items = std::sync::Arc::make_mut(items);
                        for item in items.iter_mut() {
                            *item = VBValue::Empty;
                        }
                    }
                    _ => {
                        *v = VBValue::Empty;
                    }
                }
            }
        }
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
