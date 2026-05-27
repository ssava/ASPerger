use ahash::AHashMap;

use super::VBValue;

#[derive(Default)]
pub struct ExecutionContext {
    variables: AHashMap<String, VBValue>,
    pub response_buffer: String,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: AHashMap::new(),
            response_buffer: String::new(),
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
}