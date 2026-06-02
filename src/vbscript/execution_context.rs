use ahash::AHashMap;

use super::block::UserDefinedFunction;
use super::tokenizer::Token;
use super::vbs_error::VBSError;
use super::VBValue;

#[derive(PartialEq)]
pub enum ErrorMode {
    Normal,
    ResumeNext,
}

impl Default for ErrorMode {
    fn default() -> Self {
        ErrorMode::Normal
    }
}

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
    error_mode: ErrorMode,
    pub err_number: f64,
    pub err_description: String,
    pub with_object: Option<VBValue>,

    // Request data (set before each request)
    pub request_method: String,
    pub request_path: String,
    pub request_query_string: String,
    pub request_params: AHashMap<String, String>,
    pub request_headers: AHashMap<String, String>,
    pub request_form: AHashMap<String, String>,
    pub request_cookies: AHashMap<String, String>,

    // Response control
    pub response_status: String,
    pub response_extra_headers: Vec<(String, String)>,
    pub response_ended: bool,
    pub response_redirect_url: String,

    // Session
    pub session_id: String,

    // Debugger
    pub debugger: Option<super::debugger::Debugger>,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            variables: AHashMap::new(),
            response_buffer: String::new(),
            functions: AHashMap::new(),
            classes: AHashMap::new(),
            error_mode: ErrorMode::Normal,
            err_number: 0.0,
            err_description: String::new(),
            with_object: None,
            request_method: "GET".to_string(),
            request_path: String::new(),
            request_query_string: String::new(),
            request_params: AHashMap::new(),
            request_headers: AHashMap::new(),
            request_form: AHashMap::new(),
            request_cookies: AHashMap::new(),
            response_status: "200 OK".to_string(),
            response_extra_headers: Vec::new(),
            response_ended: false,
            response_redirect_url: String::new(),
            session_id: String::new(),
            debugger: None,
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

    pub fn get_error_mode(&self) -> &ErrorMode {
        &self.error_mode
    }

    pub fn set_error_mode(&mut self, mode: ErrorMode) {
        self.error_mode = mode;
    }

    pub fn set_err(&mut self, err: VBSError) {
        self.err_number = err.code as f64;
        self.err_description = err.message;
    }

    pub fn clear_err(&mut self) {
        self.err_number = 0.0;
        self.err_description.clear();
    }

}
