use std::hash::{Hash, Hasher};
use std::sync::Arc;

use ahash::AHashMap;

use super::block::{BlockStatement, UserDefinedFunction};
use super::debugger::Debugger;
use super::store::Store;
use super::tokenizer::Token;
use super::vbs_error::VBSError;
use super::VBValue;

// ── Case-insensitive string types for hashmap keys ──────────────────────

/// Case-insensitive string slice for hashmap lookups (borrowed key).
/// Uses ASCII case-insensitive hashing and comparison.
#[repr(transparent)]
pub struct CIStr(str);

impl CIStr {
    pub fn new(s: &str) -> &Self {
        unsafe { &*(s as *const str as *const CIStr) }
    }
}

impl Hash for CIStr {
    fn hash<H: Hasher>(&self, state: &mut H) {
        for b in self.0.bytes() {
            state.write_u8(b.to_ascii_lowercase());
        }
    }
}

impl PartialEq for CIStr {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(&other.0)
    }
}
impl Eq for CIStr {}

/// Owned case-insensitive string for HashMap keys.
#[derive(Debug, Clone)]
pub struct CIString(String);

impl CIString {
    pub fn new(s: String) -> Self {
        CIString(s)
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl Hash for CIString {
    fn hash<H: Hasher>(&self, state: &mut H) {
        CIStr::new(&self.0).hash(state)
    }
}

impl PartialEq for CIString {
    fn eq(&self, other: &Self) -> bool {
        CIStr::new(&self.0) == CIStr::new(&other.0)
    }
}
impl Eq for CIString {}

impl std::borrow::Borrow<CIStr> for CIString {
    fn borrow(&self) -> &CIStr {
        CIStr::new(&self.0)
    }
}

/// VBScript error-handling mode.
///
/// - `Normal`: errors halt execution and propagate up the call stack.
/// - `ResumeNext`: errors are recorded in `Err` but execution continues
///   on the next statement.
#[derive(PartialEq, Default)]
pub enum ErrorMode {
    #[default]
    Normal,
    ResumeNext,
}

/// Parsed definition of a VBScript `Property Get/Let/Set` block.
#[allow(dead_code)]
pub struct PropertyDef {
    pub name: String,
    pub get_body: Option<Vec<Vec<Token>>>,
    pub let_body: Option<Vec<Vec<Token>>>,
    pub let_param: Option<String>,
    pub set_body: Option<Vec<Vec<Token>>>,
    pub set_param: Option<String>,
}

/// Parsed definition of a `Sub` or `Function` method.
#[allow(dead_code)]
#[derive(Clone)]
pub struct MethodDef {
    pub name: String,
    pub params: Vec<String>,
    pub body_lines: Vec<Vec<Token>>,
    pub is_function: bool,
}

/// Parsed `Class` definition with its properties and methods.
#[allow(dead_code)]
pub struct ClassDefinition {
    pub name: String,
    pub properties: AHashMap<String, PropertyDef>,
    pub methods: AHashMap<String, MethodDef>,
}

/// Variable scope: holds variables, functions, classes, error state, and the with-object.
///
/// Each `Scope` is the per-request variable store.  When a function or
/// class method is called, a new scope inherits or wraps the parent's
/// variables (see `with_instance_scope` / `with_class_method_scope`).
pub struct Scope {
    /// All script-level variables (case-insensitive keys).
    variables: AHashMap<CIString, VBValue>,
    /// User-defined `Sub` / `Function` definitions.
    functions: AHashMap<CIString, UserDefinedFunction>,
    /// Cached parsed function bodies — parsed once at definition time, reused on every call.
    function_bodies: AHashMap<CIString, Vec<BlockStatement>>,
    /// `Class` definitions (stored by class name).
    classes: AHashMap<CIString, ClassDefinition>,
    /// Current `On Error` mode.
    error_mode: ErrorMode,
    /// The `Err.Number` value set by the last runtime error.
    pub err_number: f64,
    /// The `Err.Description` value set by the last runtime error.
    pub err_description: String,
    /// The object set by `With obj ... End With`.
    pub with_object: Option<VBValue>,
    /// The expression value set by `Select Case expr`.
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
        self.variables.get(CIStr::new(name))
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(CIString::new(name.to_string()), value);
    }

    pub fn get_variable_mut(&mut self, name: &str) -> Option<&mut VBValue> {
        self.variables.get_mut(CIStr::new(name))
    }

    pub fn define_function(&mut self, func: UserDefinedFunction) {
        self.functions.insert(CIString::new(func.name.clone()), func);
    }

    pub fn get_function(&self, name: &str) -> Option<&UserDefinedFunction> {
        self.functions.get(CIStr::new(name))
    }

    pub fn get_function_body(&self, name: &str) -> Option<&Vec<BlockStatement>> {
        self.function_bodies.get(CIStr::new(name))
    }

    pub fn set_function_body(&mut self, name: &str, body: Vec<BlockStatement>) {
        self.function_bodies.insert(CIString::new(name.to_string()), body);
    }

    pub fn define_class(&mut self, class: ClassDefinition) {
        self.classes.insert(CIString::new(class.name.clone()), class);
    }

    pub fn get_class(&self, name: &str) -> Option<&ClassDefinition> {
        self.classes.get(CIStr::new(name))
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

    pub fn variables(&self) -> &AHashMap<CIString, VBValue> {
        &self.variables
    }

    pub fn variables_mut(&mut self) -> &mut AHashMap<CIString, VBValue> {
        &mut self.variables
    }
}

/// Per-request HTTP data populated by the server before script execution.
///
/// Populated by `process_request` in `server.rs` before the handler chain
/// runs.  `params` holds URL query parameters (`?a=1&b=2`), `form` holds
/// POST body (URL-encoded or multipart), and `cookies` is the parsed
/// `Cookie` header.
#[derive(Default)]
pub struct RequestContext {
    /// HTTP method (GET, POST, HEAD, etc.).
    pub method: String,
    /// URL path (without query string).
    pub path: String,
    /// Raw query string portion of the URL.
    pub query_string: String,
    /// Parsed query-string key-value pairs.
    pub params: AHashMap<String, String>,
    /// All request headers (lowercase keys).
    pub headers: AHashMap<String, String>,
    /// POST form data (URL-encoded or multipart).
    pub form: AHashMap<String, String>,
    /// Parsed Cookie header key-value pairs.
    pub cookies: AHashMap<String, String>,
    /// Content-Length (byte count of the request body).
    pub total_bytes: usize,
    /// Active code page for string encoding.
    pub code_page: u32,
    /// Locale identifier.
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
///
/// Written to by `Response.Write`, `Response.Redirect`, `Response.End`,
/// and the cookie setters.  After the handler chain finishes, the server
/// flattens this into an HTTP response line, headers, and body.
#[derive(Default)]
pub struct ResponseContext {
    /// Response body content built by `Response.Write` calls.
    pub buffer: String,
    /// HTTP status line (e.g. `"200 OK"`, `"302 Found"`).
    pub status: String,
    /// Extra headers to append to the response (e.g. `Set-Cookie`, `Location`).
    pub extra_headers: Vec<(String, String)>,
    /// Set by `Response.End` / `Response.Redirect` — the handler chain
    /// should stop processing further blocks when this is true.
    pub ended: bool,
    /// URL set by `Response.Redirect` for the `Location` header.
    pub redirect_url: String,
    /// Content flushed via `Response.Flush` (already sent to the client).
    pub flushed: String,
    /// Cookies set via `Response.Cookies("name") = value`.
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
    #[allow(clippy::type_complexity)]
    pub execute_file_callback:
        Option<Arc<dyn Fn(&str, &mut ExecutionContext) -> Result<(), String> + Send + Sync>>,
    /// Physical ASP file line where the current VBScript code block starts.
    /// Used to offset `block.line()` values to match VS Code breakpoints.
    pub code_start_line: usize,
    /// Unique per-request ID for Application.Lock ownership tracking.
    pub request_id: u64,
}

impl ExecutionContext {
    /// Create a new execution context with defaults.
    pub fn new() -> Self {
        Self::default()
    }
}

impl Default for ExecutionContext {
    fn default() -> Self {
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
            request_id: 0,
        }
    }
}

impl ExecutionContext {
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
        instance_vars: &mut AHashMap<CIString, VBValue>,
        f: impl FnOnce(&mut Self) -> Result<T, VBSError>,
    ) -> Result<T, VBSError> {
        let saved = std::mem::replace(self.scope.variables_mut(), std::mem::take(instance_vars));
        let result = f(self);
        *instance_vars = std::mem::replace(self.scope.variables_mut(), saved);
        result
    }

    /// Run closure `f` with `instance_vars` merged into the current scope
    /// (instance vars take priority over globals). After the closure, updated
    /// instance vars are extracted back and globals remain in scope.
    /// Used for class Sub/Function method dispatch.
    pub fn with_class_method_scope<T>(
        &mut self,
        instance_vars: &mut AHashMap<CIString, VBValue>,
        f: impl FnOnce(&mut Self) -> Result<T, VBSError>,
    ) -> Result<T, VBSError> {
        let saved = std::mem::take(self.scope.variables_mut());
        let mut merged = saved.clone();
        for (k, v) in instance_vars.iter() {
            merged.insert(k.clone(), v.clone());
        }
        *self.scope.variables_mut() = merged;
        let result = f(self);
        let final_vars = self.scope.variables_mut();
        for key in instance_vars.keys().cloned().collect::<Vec<_>>() {
            if let Some(v) = final_vars.remove(&key) {
                instance_vars.insert(key, v);
            }
        }
        // What's left in final_vars are the globals (possibly mutated) — keep them
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
