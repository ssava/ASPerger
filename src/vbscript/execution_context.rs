use std::collections::HashMap;

use super::VBValue;

#[derive(Default)]
pub struct ExecutionContext {
    variables: HashMap<String, VBValue>,
    pub response_buffer: String,
    functions: HashMap<String, VBValue>, // Store functions
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: HashMap::new(),
            response_buffer: String::new(),
            functions: HashMap::new(),
        }
    }

    pub fn flush_response_buffer(&mut self) {
        self.response_buffer.clear();
    }

    pub fn write(&mut self, content: &str) {
        self.response_buffer.push_str(content);
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(name.to_string(), value);
    }

    pub fn get_variable(&self, name: &str) -> Option<VBValue> {
        self.variables.get(name).cloned()
    }

    pub fn set_function(&mut self, name: String, params: Vec<String>, body: String) {
        self.functions.insert(name, VBValue::Function(params, body));
    }

    pub fn get_function(&self, name: &str) -> Option<&VBValue> {
        self.functions.get(name)
    }
}