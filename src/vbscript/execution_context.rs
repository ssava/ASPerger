use std::sync::Arc;

use ahash::AHashMap;

use super::block::{BlockStatement, UserDefinedFunction};
use super::debugger::Debugger;
use super::store::Store;
use super::tokenizer::Token;
use super::vbs_error::VBSError;
use super::VBValue;

type ExecuteFileCallback =
    Arc<dyn Fn(&str, &mut ExecutionContext) -> Result<(), String> + Send + Sync>;

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
pub struct PropertyDef {
    pub name: String,
    pub get_body: Option<Vec<Vec<Token>>>,
    pub let_body: Option<Vec<Vec<Token>>>,
    pub let_param: Option<String>,
}

/// Parsed definition of a `Sub` or `Function` method.
#[derive(Clone)]
pub struct MethodDef {
    pub name: String,
    pub params: Vec<String>,
    pub body_lines: Vec<Vec<Token>>,
    pub is_function: bool,
}

/// Parsed `Class` definition with its properties and methods.
pub struct ClassDefinition {
    pub name: String,
    pub properties: AHashMap<String, PropertyDef>,
    pub methods: AHashMap<String, MethodDef>,
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
pub struct ExecutionContext {
    /// All script-level variables (case-insensitive keys).
    variables: AHashMap<String, VBValue>,
    /// User-defined `Sub` / `Function` definitions.
    functions: AHashMap<String, UserDefinedFunction>,
    /// Cached parsed function bodies.
    function_bodies: AHashMap<String, Vec<BlockStatement>>,
    /// `Class` definitions (stored by class name).
    classes: AHashMap<String, ClassDefinition>,
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
    pub execute_file_callback: Option<ExecuteFileCallback>,
    /// Physical ASP file line where the current VBScript code block starts.
    pub code_start_line: usize,
    /// Unique per-request ID for Application.Lock ownership tracking.
    pub request_id: u64,
}

/// Scoped guard that resets `code_start_line` to a saved value on drop.
///
/// Used in `execute_user_defined_function` to ensure the field is restored
/// even on early returns or panics.
pub(crate) struct CodeStartLineGuard {
    code_start_line: *mut usize,
    saved: usize,
}

impl CodeStartLineGuard {
    pub fn new(code_start_line: &mut usize) -> Self {
        let saved = *code_start_line;
        *code_start_line = 0;
        CodeStartLineGuard {
            code_start_line,
            saved,
        }
    }
}

impl Drop for CodeStartLineGuard {
    fn drop(&mut self) {
        unsafe { *self.code_start_line = self.saved; }
    }
}

impl ExecutionContext {
    pub fn new() -> Self {
        Self::default()
    }

    fn lc_key<'a>(&self, name: &'a str) -> std::borrow::Cow<'a, str> {
        if name.bytes().any(|b| b.is_ascii_uppercase()) {
            name.to_lowercase().into()
        } else {
            name.into()
        }
    }

    pub fn get_variable(&self, name: &str) -> Option<&VBValue> {
        self.variables.get(self.lc_key(name).as_ref())
    }

    pub fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(name.to_lowercase(), value);
    }

    pub fn get_variable_mut(&mut self, name: &str) -> Option<&mut VBValue> {
        self.variables.get_mut(self.lc_key(name).as_ref())
    }

    pub fn define_function(&mut self, func: UserDefinedFunction) {
        self.functions.insert(func.name.to_lowercase(), func);
    }

    pub fn get_function(&self, name: &str) -> Option<&UserDefinedFunction> {
        self.functions.get(self.lc_key(name).as_ref())
    }

    pub fn get_function_body(&self, name: &str) -> Option<&Vec<BlockStatement>> {
        self.function_bodies.get(self.lc_key(name).as_ref())
    }

    pub fn set_function_body(&mut self, name: &str, body: Vec<BlockStatement>) {
        self.function_bodies.insert(name.to_lowercase(), body);
    }

    pub fn define_class(&mut self, class: ClassDefinition) {
        self.classes.insert(class.name.to_lowercase(), class);
    }

    pub fn get_class(&self, name: &str) -> Option<&ClassDefinition> {
        self.classes.get(self.lc_key(name).as_ref())
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

    /// Clear the response buffer.
    pub fn flush_response_buffer(&mut self) {
        self.response.flush_buffer();
    }

    /// Write a string to the response buffer.
    pub fn write(&mut self, content: &str) {
        self.response.write(content);
    }

    /// Temporarily replace the current variables with `instance_vars`,
    /// run closure `f`, then restore. Used for class Property Get/Let/Set.
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

    /// Run closure `f` with `instance_vars` merged into variables (instance vars
    /// take priority). After closure, updated instance vars are extracted back.
    pub fn with_class_method_scope<T>(
        &mut self,
        instance_vars: &mut AHashMap<String, VBValue>,
        f: impl FnOnce(&mut Self) -> Result<T, VBSError>,
    ) -> Result<T, VBSError> {
        let saved = std::mem::take(&mut self.variables);
        let mut merged = saved.clone();
        for (k, v) in instance_vars.iter() {
            merged.insert(k.clone(), v.clone());
        }
        self.variables = merged;
        let result = f(self);
        for key in instance_vars.keys().cloned().collect::<Vec<_>>() {
            if let Some(v) = self.variables.remove(&key) {
                instance_vars.insert(key, v);
            }
        }
        result
    }
}

impl Default for ExecutionContext {
    fn default() -> Self {
        ExecutionContext {
            variables: AHashMap::new(),
            functions: AHashMap::new(),
            function_bodies: AHashMap::new(),
            classes: AHashMap::new(),
            error_mode: ErrorMode::Normal,
            err_number: 0.0,
            err_description: String::new(),
            with_object: None,
            select_value: None,
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
