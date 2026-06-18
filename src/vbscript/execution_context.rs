use std::sync::Arc;

use ahash::AHashMap;

use super::block::{BlockStatement, UserDefinedFunction};
use super::debugger::Debugger;
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

/// Variable scope: holds variables, functions, classes, error state, and the with-object.
pub struct Scope {
    variables: AHashMap<String, VBValue>,
    functions: AHashMap<String, UserDefinedFunction>,
    /// Cached parsed function bodies — parsed once at definition time, reused on every call.
    function_bodies: AHashMap<String, Vec<BlockStatement>>,
    classes: AHashMap<String, ClassDefinition>,
    error_mode: ErrorMode,
    pub err_number: f64,
    pub err_description: String,
    pub with_object: Option<VBValue>,
    pub(crate) select_value: Option<VBValue>,
}

impl Scope {
    fn new() -> Self {
        Scope {
            variables: AHashMap::new(),
            functions: AHashMap::new(),
            function_bodies: AHashMap::new(),
            classes: AHashMap::new(),
            error_mode: ErrorMode::Normal,
            err_number: 0.0,
            err_description: String::new(),
            with_object: None,
            select_value: None,
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

    pub fn get_function_body(&self, name: &str) -> Option<&Vec<BlockStatement>> {
        self.function_bodies.get(&name.to_uppercase())
    }

    pub fn set_function_body(&mut self, name: &str, body: Vec<BlockStatement>) {
        self.function_bodies.insert(name.to_uppercase(), body);
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

/// Per-request HTTP data populated by the server before script execution.
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

/// Per-cookie data stored during Response.Cookies set operations.
#[derive(Debug, Clone, Default)]
pub struct CookieEntry {
    pub value: String,
    pub subkeys: AHashMap<String, String>,
    pub expires: String,
    pub domain: String,
    pub path: String,
    pub secure: bool,
}

/// Response state accumulated during script execution.
#[derive(Default)]
pub struct ResponseContext {
    pub buffer: String,
    pub status: String,
    pub extra_headers: Vec<(String, String)>,
    pub ended: bool,
    pub redirect_url: String,
    pub flushed: String,
    pub cookies: AHashMap<String, CookieEntry>,
}

impl ResponseContext {
    pub fn write(&mut self, content: &str) {
        self.buffer.push_str(content);
    }

    pub fn flush_buffer(&mut self) {
        self.buffer.clear();
    }
}

/// Session state: unique identifier and enabled flag.
#[derive(Default)]
pub struct SessionContext {
    pub id: String,
    pub enabled: bool,
}

/// Aggregate execution context that owns all per-request state.
///
/// Provides delegation methods to inner sub-contexts (`scope`, `request`,
/// `response`, `session`) and holds the shared `store` for persistence.
pub struct ExecutionContext {
    /// Variable scope, functions, classes, error state, with-object.
    pub scope: Scope,
    /// Incoming request data.
    pub request: RequestContext,
    /// Output buffer, status, headers, redirect state.
    pub response: ResponseContext,
    /// Session identifier and enabled flag.
    pub session: SessionContext,
    /// Shared session/application store (injected by the server).
    pub store: Option<Arc<Store>>,
    /// Path to the script being executed (for debugger file/breakpoint matching).
    pub script_path: String,
    /// Optional DAP debugger (shared across requests via Arc).
    pub debugger: Option<Arc<Debugger>>,
    /// Callback for Server.Execute / Server.Transfer.
    pub execute_file_callback:
        Option<Arc<dyn Fn(&str, &mut ExecutionContext) -> Result<(), String> + Send + Sync>>,
    /// Physical ASP file line where the current VBScript code block starts.
    /// Used to offset `block.line()` values to match VS Code breakpoints.
    pub code_start_line: usize,
}

impl ExecutionContext {
    /// Create a new execution context with defaults.
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
            script_path: String::new(),
            debugger: None,
            execute_file_callback: None,
            code_start_line: 0,
        }
    }

    /// Clear the response buffer.
    pub fn flush_response_buffer(&mut self) {
        self.response.flush_buffer();
    }

    /// Write a string to the response buffer.
    pub fn write(&mut self, content: &str) {
        self.response.write(content);
    }

    /// Set a variable in the current scope.
    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.scope.set_variable(name, value);
    }

    /// Get a reference to a variable in the current scope.
    pub fn get_variable(&self, name: &str) -> Option<&VBValue> {
        self.scope.get_variable(name)
    }

    /// Get a mutable reference to a variable in the current scope.
    pub fn get_variable_mut(&mut self, name: &str) -> Option<&mut VBValue> {
        self.scope.get_variable_mut(name)
    }

    /// Define a user-defined function in the current scope.
    pub fn define_function(&mut self, func: UserDefinedFunction) {
        self.scope.define_function(func);
    }

    /// Look up a user-defined function by name.
    pub fn get_function(&self, name: &str) -> Option<&UserDefinedFunction> {
        self.scope.get_function(name)
    }

    /// Look up a cached function body by function name.
    pub fn get_function_body(&self, name: &str) -> Option<&Vec<BlockStatement>> {
        self.scope.get_function_body(name)
    }

    /// Store a cached function body.
    pub fn set_function_body(&mut self, name: &str, body: Vec<BlockStatement>) {
        self.scope.set_function_body(name, body);
    }

    /// Define a class in the current scope.
    pub fn define_class(&mut self, class: ClassDefinition) {
        self.scope.define_class(class);
    }

    /// Look up a class definition by name.
    pub fn get_class(&self, name: &str) -> Option<&ClassDefinition> {
        self.scope.get_class(name)
    }

    /// Temporarily replace the current variable scope with `instance_vars`,
    /// run closure `f`, then restore the saved scope. Used for class Property Get/Let/Set.
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

    /// Get the current error mode.
    pub fn get_error_mode(&self) -> &ErrorMode {
        self.scope.get_error_mode()
    }

    /// Set the error mode (Normal or ResumeNext).
    pub fn set_error_mode(&mut self, mode: ErrorMode) {
        self.scope.set_error_mode(mode);
    }

    /// Record an error in the scope (sets err_number and err_description).
    pub fn set_err(&mut self, err: VBSError) {
        self.scope.set_err(err);
    }

    /// Clear the recorded error state.
    pub fn clear_err(&mut self) {
        self.scope.clear_err();
    }
}
