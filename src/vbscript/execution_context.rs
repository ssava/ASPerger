use ahash::AHashMap;

use super::block::UserDefinedFunction;
use super::tokenizer::Token;
use super::vbs_error::VBSError;
use super::VBValue;

#[allow(dead_code)]
pub struct PropertyDef {
    pub name: String,
    pub get_body: Option<Vec<Vec<Token>>>,
    pub let_body: Option<Vec<Vec<Token>>>,
    pub let_param: Option<String>,
    pub set_body: Option<Vec<Vec<Token>>>,
    pub set_param: Option<String>,
}

#[allow(dead_code)]
pub struct ClassDefinition {
    pub name: String,
    pub properties: AHashMap<String, PropertyDef>,
}

#[derive(Default)]
pub struct ExecutionContext {
    variables: AHashMap<String, VBValue>,
    pub response_buffer: String,
    functions: AHashMap<String, UserDefinedFunction>,
    classes: AHashMap<String, ClassDefinition>,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: AHashMap::new(),
            response_buffer: String::new(),
            functions: AHashMap::new(),
            classes: AHashMap::new(),
        }
    }

    pub fn flush_response_buffer(&mut self) {
        self.response_buffer.clear();
    }

    pub fn write(&mut self, content: &str) {
        self.response_buffer.push_str(content);
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(name.to_uppercase(), value);
    }

    pub fn get_variable(&self, name: &str) -> Option<&VBValue> {
        self.variables.get(&name.to_uppercase())
    }

    pub fn get_variable_mut(&mut self, name: &str) -> Option<&mut VBValue> {
        self.variables.get_mut(&name.to_uppercase())
    }

    pub fn define_function(&mut self, func: UserDefinedFunction) {
        self.functions.insert(func.name.to_uppercase(), func);
    }

    pub fn get_function(&self, name: &str) -> Option<&UserDefinedFunction> {
        self.functions.get(&name.to_uppercase())
    }

    pub fn define_class(&mut self, class: ClassDefinition) {
        self.classes.insert(class.name.to_uppercase(), class);
    }

    pub fn get_class(&self, name: &str) -> Option<&ClassDefinition> {
        self.classes.get(&name.to_uppercase())
    }

    pub fn with_instance_scope<T>(
        &mut self,
        instance_vars: &mut AHashMap<String, VBValue>,
        f: impl FnOnce(&mut Self) -> Result<T, VBSError>,
    ) -> Result<T, VBSError> {
        let saved = std::mem::replace(&mut self.variables, std::mem::take(instance_vars));
        let result = f(self);
        *instance_vars = std::mem::replace(&mut self.variables, saved);
        result
    }
}
