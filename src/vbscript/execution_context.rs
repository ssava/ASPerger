use std::collections::HashMap;
use super::value::VBValue;

#[derive(Default)]
pub struct ExecutionContext {
    variables: HashMap<String, VBValue>,
    pub response_buffer: String,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: HashMap::new(),
            response_buffer: String::new(),
        }
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
}