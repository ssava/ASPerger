use ahash::AHashMap;

use super::block::UserDefinedFunction;
use super::VBValue;

#[derive(Default)]
pub struct ExecutionContext {
    variables: AHashMap<String, VBValue>,
    pub response_buffer: String,
    functions: AHashMap<String, UserDefinedFunction>,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: AHashMap::new(),
            response_buffer: String::new(),
            functions: AHashMap::new(),
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
}