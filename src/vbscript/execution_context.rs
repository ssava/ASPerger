use std::sync::Arc;

use ahash::AHashMap;

use super::block::UserDefinedFunction;
use super::store::Store;
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

pub struct Scope {
    variables: AHashMap<String, VBValue>,
    functions: AHashMap<String, UserDefinedFunction>,
    classes: AHashMap<String, ClassDefinition>,
    error_mode: ErrorMode,
    pub err_number: f64,
    pub err_description: String,
    pub with_object: Option<VBValue>,
}

impl Scope {
    fn new() -> Self {
        Scope {
            variables: AHashMap::new(),
            functions: AHashMap::new(),
            classes: AHashMap::new(),
            error_mode: ErrorMode::Normal,
            err_number: 0.0,
            err_description: String::new(),
            with_object: None,
        }
    }

    pub fn get_variable(&self, name: &str) -> Option<&VBValue> {
        self.variables.get(&name.to_uppercase())
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(name.to_uppercase(), value);
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

    pub fn variables(&self) -> &AHashMap<String, VBValue> {
        &self.variables
    }

    pub fn variables_mut(&mut self) -> &mut AHashMap<String, VBValue> {
        &mut self.variables
    }
}

#[derive(Default)]
pub struct RequestContext {
    pub method: String,
    pub path: String,
    pub query_string: String,
    pub params: AHashMap<String, String>,
    pub headers: AHashMap<String, String>,
    pub form: AHashMap<String, String>,
    pub cookies: AHashMap<String, String>,
    pub total_bytes: usize,
    pub code_page: u32,
    pub lcid: u32,
}

#[derive(Default)]
pub struct ResponseContext {
    pub buffer: String,
    pub status: String,
    pub extra_headers: Vec<(String, String)>,
    pub ended: bool,
    pub redirect_url: String,
    pub flushed: String,
    pub cookies: AHashMap<String, String>,
}

impl ResponseContext {
    pub fn write(&mut self, content: &str) {
        self.buffer.push_str(content);
    }

    pub fn flush_buffer(&mut self) {
        self.buffer.clear();
    }
}

#[derive(Default)]
pub struct SessionContext {
    pub id: String,
    pub enabled: bool,
}

pub struct ExecutionContext {
    pub scope: Scope,
    pub request: RequestContext,
    pub response: ResponseContext,
    pub session: SessionContext,
    pub store: Option<Arc<Store>>,
    pub debugger: Option<super::debugger::Debugger>,
    pub execute_file_callback:
        Option<Arc<dyn Fn(&str, &mut ExecutionContext) -> Result<(), String> + Send + Sync>>,
}

impl ExecutionContext {
    pub fn new() -> Self {
        ExecutionContext {
            scope: Scope::new(),
            request: RequestContext {
                method: "GET".to_string(),
                code_page: 65001,
                lcid: 1033,
                ..RequestContext::default()
            },
            response: ResponseContext {
                status: "200 OK".to_string(),
                ..ResponseContext::default()
            },
            session: SessionContext {
                enabled: true,
                ..SessionContext::default()
            },
            store: None,
            debugger: None,
            execute_file_callback: None,
        }
    }

    pub fn flush_response_buffer(&mut self) {
        self.response.flush_buffer();
    }

    pub fn write(&mut self, content: &str) {
        self.response.write(content);
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.scope.set_variable(name, value);
    }

    pub fn get_variable(&self, name: &str) -> Option<&VBValue> {
        self.scope.get_variable(name)
    }

    pub fn get_variable_mut(&mut self, name: &str) -> Option<&mut VBValue> {
        self.scope.get_variable_mut(name)
    }

    pub fn define_function(&mut self, func: UserDefinedFunction) {
        self.scope.define_function(func);
    }

    pub fn get_function(&self, name: &str) -> Option<&UserDefinedFunction> {
        self.scope.get_function(name)
    }

    pub fn define_class(&mut self, class: ClassDefinition) {
        self.scope.define_class(class);
    }

    pub fn get_class(&self, name: &str) -> Option<&ClassDefinition> {
        self.scope.get_class(name)
    }

    pub fn with_instance_scope<T>(
        &mut self,
        instance_vars: &mut AHashMap<String, VBValue>,
        f: impl FnOnce(&mut Self) -> Result<T, VBSError>,
    ) -> Result<T, VBSError> {
        let saved = std::mem::replace(self.scope.variables_mut(), std::mem::take(instance_vars));
        let result = f(self);
        *instance_vars = std::mem::replace(self.scope.variables_mut(), saved);
        result
    }

    pub fn get_error_mode(&self) -> &ErrorMode {
        self.scope.get_error_mode()
    }

    pub fn set_error_mode(&mut self, mode: ErrorMode) {
        self.scope.set_error_mode(mode);
    }

    pub fn set_err(&mut self, err: VBSError) {
        self.scope.set_err(err);
    }

    pub fn clear_err(&mut self) {
        self.scope.clear_err();
    }
}
