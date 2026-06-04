use crate::asp::asp_error::ASPError;
use crate::asp::config::{AspServerConfig, Config};
use crate::asp::handler::{CodeHandler, Handler, HtmlHandler};
use crate::asp::parser::AspParser;
use crate::vbscript::debugger::Debugger;
use crate::vbscript::{store::Store, ExecutionContext, Interpreter, VBScriptInterpreter, VBValue};
use ahash::AHashMap;
use std::path::Path;
use std::sync::Arc;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::net::TcpListener;

/// Parsed HTTP request received by the server.
#[derive(Debug, Clone)]
pub struct HttpRequest {
    /// HTTP method (GET, POST, etc.).
    pub method: String,
    /// Request path (e.g. "index.asp").
    pub path: String,
    /// Raw query string portion of the URL.
    pub query_string: String,
    /// Request headers keyed by lowercased name.
    pub headers: AHashMap<String, String>,
    /// Raw request body bytes.
    pub body: Vec<u8>,
    /// Parsed cookies from the Cookie header.
    pub cookies: AHashMap<String, String>,
}

/// HTTP response to be written to the client.
#[derive(Debug, Clone)]
pub struct HttpResponse {
    /// Status line (e.g. "200 OK", "404 Not Found").
    pub status_line: String,
    /// Content-Type header value.
    pub content_type: String,
    /// Response body string.
    pub body: String,
    /// Additional headers to include in the response.
    pub extra_headers: Vec<(String, String)>,
}

/// Main ASP server, owning the handler chain, shared store, and config.
pub struct AspServer {
    /// Chain of handlers that process ASP blocks.
    pub handler_chain: Arc<dyn Handler + Send + Sync>,
    /// Shared session/application data store.
    pub store: Arc<Store>,
    /// Server configuration (host, port, folder).
    config: Config,
}

impl AspServer {
    /// Create a new `AspServer` with the given configuration.
    /// Initializes the VBScript interpreter, sets up the handler chain
    /// (HtmlHandler → CodeHandler), and creates a shared `Store`.
    pub fn new(config: Config) -> Self {
        let interpreter: Arc<dyn Interpreter> = Arc::new(VBScriptInterpreter);

        let mut html_handler = HtmlHandler::new();
        let code_handler = CodeHandler::new(Arc::clone(&interpreter));

        html_handler.set_next(Arc::new(code_handler));

        AspServer {
            handler_chain: Arc::new(html_handler),
            store: Store::new(),
            config,
        }
    }

    /// Start the HTTP server, listening on the configured host:port.
    /// Each incoming connection is handled in a separate Tokio task.
    pub async fn start(&self) -> std::io::Result<()> {
        self.start_with_config(&Default::default()).await
    }

    /// Start the HTTP server using a full `AspServerConfig` (includes ini file support).
    pub async fn start_with_config(&self, asp_cfg: &AspServerConfig) -> std::io::Result<()> {
        let host = if asp_cfg.host.is_empty() { &self.config.host } else { &asp_cfg.host };
        let port = if asp_cfg.port == 0 { self.config.port } else { asp_cfg.port };
        let folder = if asp_cfg.folder.is_empty() || asp_cfg.folder == "." {
            self.config.folder.clone()
        } else {
            asp_cfg.folder.clone()
        };
        let default_doc = asp_cfg.default_document.clone();

        let bind_addr = format!("{}:{}", host, port);
        let listener = TcpListener::bind(&bind_addr).await?;
        println!(
            "Server listening on {} serving files from {} (default document: {})",
            bind_addr, folder, default_doc
        );

        loop {
            let (mut stream, _) = listener.accept().await?;
            let handler_chain = Arc::clone(&self.handler_chain);
            let store = Arc::clone(&self.store);
            let folder = folder.clone();
            let default_doc = default_doc.clone();

            tokio::spawn(async move {
                if let Err(e) =
                    Self::handle_connection(&handler_chain, &mut stream, &folder, &default_doc, &store).await
                {
                    eprintln!("Connection handling error: {}", e);
                }
            });
        }
    }

    fn url_decode(s: &str) -> String {
        let mut result = String::with_capacity(s.len());
        let mut chars = s.bytes();
        while let Some(b) = chars.next() {
            match b {
                b'+' => result.push(' '),
                b'%' => {
                    let hi = chars
                        .next()
                        .and_then(|c| (c as char).to_digit(16))
                        .unwrap_or(0);
                    let lo = chars
                        .next()
                        .and_then(|c| (c as char).to_digit(16))
                        .unwrap_or(0);
                    result.push((hi * 16 + lo) as u8 as char);
                }
                _ => result.push(b as char),
            }
        }
        result
    }

    fn parse_query_string(query: &str) -> AHashMap<String, String> {
        let mut params = AHashMap::new();
        if query.is_empty() {
            return params;
        }
        for pair in query.split('&') {
            if let Some((key, value)) = pair.split_once('=') {
                let decoded_key = Self::url_decode(key);
                let decoded_value = Self::url_decode(value);
                params.insert(decoded_key, decoded_value);
            } else if !pair.is_empty() {
                params.insert(Self::url_decode(pair), String::new());
            }
        }
        params
    }

    fn parse_cookies(cookie_header: &str) -> AHashMap<String, String> {
        let mut cookies = AHashMap::new();
        for pair in cookie_header.split(';') {
            let pair = pair.trim();
            if let Some((key, value)) = pair.split_once('=') {
                cookies.insert(key.trim().to_string(), value.trim().to_string());
            }
        }
        cookies
    }

    fn rand_hex() -> String {
        let val: u64 = {
            use std::time::{SystemTime, UNIX_EPOCH};
            let d = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default();
            (d.as_nanos() & 0xFFFFFFFF) as u64
        };
        format!("{:04x}", val)
    }

    fn parse_multipart_form_data(body: &[u8], boundary: &str) -> AHashMap<String, String> {
        let mut form = AHashMap::new();
        let body_str = String::from_utf8_lossy(body);

        let full_boundary = format!("--{}", boundary);
        for part in body_str.split(&full_boundary) {
            let part = part.trim();
            if part.is_empty() || part.starts_with("--") || !part.contains("\r\n\r\n") {
                continue;
            }
            if let Some((headers, content)) = part.split_once("\r\n\r\n") {
                let content = content.trim_end_matches("\r\n").trim_end();
                if let Some(name_start) = headers.find("name=\"") {
                    let name_start = name_start + 6;
                    if let Some(name_end) = headers[name_start..].find('"') {
                        let name = &headers[name_start..name_start + name_end];
                        form.insert(name.to_string(), content.to_string());
                    }
                }
            }
        }
        form
    }

    fn parse_form_body(body: &[u8]) -> AHashMap<String, String> {
        let body_str = String::from_utf8_lossy(body);
        let mut form = AHashMap::new();
        for pair in body_str.split('&') {
            if let Some((key, value)) = pair.split_once('=') {
                let decoded_key = Self::url_decode(key);
                let decoded_value = Self::url_decode(value);
                form.insert(decoded_key, decoded_value);
            } else if !pair.is_empty() {
                form.insert(Self::url_decode(pair), String::new());
            }
        }
        form
    }

    fn generate_session_id() -> String {
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .subsec_nanos();
        format!("ASPERGER{:x}{}", nanos, Self::rand_hex())
    }

    pub async fn write_response(
        stream: &mut tokio::net::TcpStream,
        response: &HttpResponse,
    ) -> Result<(), ASPError> {
        let has_content_type = response
            .extra_headers
            .iter()
            .any(|(k, _)| k.eq_ignore_ascii_case("content-type"));
        let mut buf = format!(
            "HTTP/1.1 {}\r\nContent-Length: {}\r\n",
            response.status_line,
            response.body.len(),
        );
        if !has_content_type {
            buf.push_str(&format!(
                "Content-Type: {}; charset=utf-8\r\n",
                response.content_type
            ));
        }
        for (key, value) in &response.extra_headers {
            buf.push_str(&format!("{}: {}\r\n", key, value));
        }
        buf.push_str("\r\n");
        buf.push_str(&response.body);

        stream
            .write_all(buf.as_bytes())
            .await
            .map_err(|e| ASPError::new(500, format!("Error writing response: {}", e)))?;
        stream
            .flush()
            .await
            .map_err(|e| ASPError::new(500, format!("Error flushing buffer: {}", e)))?;
        Ok(())
    }

    /// Read and parse an HTTP request from the given stream.
    /// Returns an `HttpRequest` with method, path, headers, body, and cookies.
    pub async fn read_request(stream: &mut tokio::net::TcpStream) -> Result<HttpRequest, ASPError> {
        use tokio::io::AsyncBufReadExt;
        use tokio::io::BufReader;

        let mut reader = BufReader::new(&mut *stream);
        let mut headers = AHashMap::new();

        let mut request_line = String::new();
        reader
            .read_line(&mut request_line)
            .await
            .map_err(|e| ASPError::new(500, format!("Error reading request line: {}", e)))?;
        let mut parts = request_line.split_whitespace();
        let method = parts.next().unwrap_or("GET").to_string();
        let full_path = parts.next().unwrap_or("/").to_string();
        let (path, query_string) = match full_path.split_once('?') {
            Some((p, q)) => (p.trim_start_matches('/').to_string(), q.to_string()),
            None => (full_path.trim_start_matches('/').to_string(), String::new()),
        };

        loop {
            let mut line = String::new();
            let n = reader
                .read_line(&mut line)
                .await
                .map_err(|e| ASPError::new(500, format!("Error reading header: {}", e)))?;
            if n <= 1 || line.trim().is_empty() {
                break;
            }
            if let Some((key, value)) = line.split_once(':') {
                headers.insert(key.trim().to_lowercase(), value.trim().to_string());
            }
        }

        let cookies = headers
            .get("cookie")
            .map(|c| Self::parse_cookies(c))
            .unwrap_or_default();

        let content_length: usize = headers
            .get("content-length")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        let body = if content_length > 0 {
            let mut body = vec![0u8; content_length];
            reader
                .read_exact(&mut body)
                .await
                .map_err(|e| ASPError::new(500, format!("Error reading body: {}", e)))?;
            body
        } else {
            Vec::new()
        };

        Ok(HttpRequest {
            method,
            path,
            query_string,
            headers,
            body,
            cookies,
        })
    }

    /// Inject the five ASP intrinsic objects (Request, Response, Session,
    /// Server, Application) into the execution context as global variables.
    /// Safe to call multiple times (only injects missing objects).
    pub fn inject_asp_intrinsic_objects(context: &mut ExecutionContext) {
        use crate::vbscript::asp_objects::*;
        if context.get_variable("REQUEST").is_none() {
            context.set_variable("Request", VBValue::Object(Box::new(RequestObject)));
        }
        if context.get_variable("RESPONSE").is_none() {
            context.set_variable("Response", VBValue::Object(Box::new(ResponseObject)));
        }
        if context.get_variable("SESSION").is_none() {
            let session = SessionObject {
                session_id: context.session.id.clone(),
                session_enabled: context.session.enabled,
            };
            context.set_variable("Session", VBValue::Object(Box::new(session)));
        }
        if context.get_variable("SERVER").is_none() {
            context.set_variable("Server", VBValue::Object(Box::new(ServerObject)));
        }
        if context.get_variable("APPLICATION").is_none() {
            context.set_variable("Application", VBValue::Object(Box::new(ApplicationObject)));
        }
    }

    /// Process a parsed HTTP request: resolve includes, parse ASP blocks,
    /// apply preprocessor directives, inject intrinsic objects, and execute
    /// blocks through the handler chain. Returns an `HttpResponse`.
    pub async fn process_request(
        request: HttpRequest,
        handler_chain: &Arc<dyn Handler + Send + Sync>,
        folder: &str,
        default_document: &str,
        store: &Arc<Store>,
        debugger: Option<Arc<Debugger>>,
    ) -> Result<HttpResponse, ASPError> {
        let file_path = format!(
            "{}/{}",
            folder,
            if request.path.is_empty() {
                default_document.to_string()
            } else {
                request.path.clone()
            }
        );
        let path_obj = Path::new(&file_path);

        eprintln!("[ASP] file_path={:?}", file_path);

        let canonical_path = match path_obj.canonicalize() {
            Ok(p) => {
                eprintln!("[ASP] canonical_path={:?}", p);
                p
            }
            Err(e) => {
                let err = ASPError::new(404, &format!("File not found: {} (folder={}, path={}, error={})", file_path, folder, request.path, e));
                eprintln!("[ASP] canonicalize failed: {}", e);
                return Ok(HttpResponse {
                    status_line: "404 Not Found".to_string(),
                    content_type: "text/html".to_string(),
                    body: err.render_html(),
                    extra_headers: Vec::new(),
                });
            }
        };

        let canonical_folder = Path::new(folder)
            .canonicalize()
            .map_err(|_| ASPError::new(500, "Server configuration error".to_string()))?;
        eprintln!("[ASP] canonical_folder={:?}", canonical_folder);

        if !canonical_path.starts_with(canonical_folder) {
            let err = ASPError::new(403, "Forbidden: access denied");
            return Ok(HttpResponse {
                status_line: "403 Forbidden".to_string(),
                content_type: "text/html".to_string(),
                body: err.render_html(),
                extra_headers: Vec::new(),
            });
        }

        let content = match std::fs::read_to_string(&file_path) {
            Ok(content) => content,
            Err(_) => {
                return Ok(HttpResponse {
                    status_line: "404 Not Found".to_string(),
                    content_type: "text/plain".to_string(),
                    body: format!("Page not found: {}", file_path),
                    extra_headers: Vec::new(),
                });
            }
        };

        if !file_path.ends_with(".asp") {
            let content_type = match path_obj.extension().and_then(|e| e.to_str()) {
                Some("html") | Some("htm") => "text/html",
                Some("css") => "text/css",
                Some("js") => "application/javascript",
                Some("txt") => "text/plain",
                _ => "application/octet-stream",
            };
            return Ok(HttpResponse {
                status_line: "200 OK".to_string(),
                content_type: content_type.to_string(),
                body: content,
                extra_headers: Vec::new(),
            });
        }

        // ASP file processing
        let file_dir = Path::new(&file_path).parent().unwrap_or(Path::new(folder));
        let root_dir = Path::new(folder);

        let expanded = match crate::asp::include_resolver::IncludeResolver::expand(
            &content, file_dir, root_dir,
        ) {
            Ok(s) => s,
            Err(e) => {
                return Ok(HttpResponse {
                    status_line: "500 Internal Server Error".to_string(),
                    content_type: "text/html".to_string(),
                    body: ASPError::new(500, e).render_html(),
                    extra_headers: Vec::new(),
                });
            }
        };

        let parser = AspParser::new(expanded);
        let blocks = parser.parse();

        let preprocessor = crate::asp::preprocessor::Preprocessor::new();
        let (directive_config, filtered_blocks) = preprocessor.process(&blocks);

        let mut context = ExecutionContext::new();
        context.script_path = file_path.clone();
        context.store = Some(Arc::clone(store));
        context.session.enabled = directive_config.enable_session_state;
        if let Some(cp) = directive_config.code_page {
            context.request.code_page = cp;
        }
        if let Some(l) = directive_config.lcid {
            context.request.lcid = l;
        }

        context.request.method = request.method.clone();
        context.request.path = request.path.clone();
        context.request.query_string = request.query_string.clone();
        context.request.params = Self::parse_query_string(&request.query_string);
        context.request.headers = request.headers.clone();
        context.request.cookies = request.cookies;
        context.request.total_bytes = request.body.len();

        let content_type = request
            .headers
            .get("content-type")
            .cloned()
            .unwrap_or_default();
        if request.method.eq_ignore_ascii_case("POST") {
            if content_type.contains("application/x-www-form-urlencoded") {
                context.request.form = Self::parse_form_body(&request.body);
            } else if content_type.contains("multipart/form-data") {
                if let Some(boundary) = content_type
                    .split(';')
                    .find_map(|p| p.trim().strip_prefix("boundary="))
                {
                    context.request.form = Self::parse_multipart_form_data(&request.body, boundary);
                }
            }
        }

        // Session handling (only when enabled)
        if context.session.enabled {
            let existing_session = context
                .request
                .cookies
                .get("ASPSESSIONID")
                .cloned()
                .unwrap_or_default();
            let session_was_new = existing_session.is_empty();
            if session_was_new {
                context.session.id = Self::generate_session_id();
            } else {
                context.session.id = existing_session;
            }

            if session_was_new && !context.session.id.is_empty() {
                context.response.extra_headers.push((
                    "Set-Cookie".to_string(),
                    format!("ASPSESSIONID={}; path=/", context.session.id),
                ));
            }
        }

        // Set up Server.Execute/Transfer callback
        let folder_clone = folder.to_string();
        let handler_clone = Arc::clone(handler_chain);
        context.execute_file_callback = Some(Arc::new(move |path, ctx| {
            let target = if path.starts_with('/') || path.starts_with('\\') {
                format!("{}{}", folder_clone, path)
            } else {
                format!("{}/{}", folder_clone, path)
            };
            let content = std::fs::read_to_string(&target)
                .map_err(|e| format!("Could not read '{}': {}", target, e))?;
            let target_dir = Path::new(&target)
                .parent()
                .unwrap_or(Path::new(&folder_clone));
            let root = Path::new(&folder_clone);
            let expanded =
                crate::asp::include_resolver::IncludeResolver::expand(&content, target_dir, root)
                    .map_err(|e| format!("Include error in '{}': {}", target, e))?;
            let p = crate::asp::parser::AspParser::new(expanded);
            let inner_blocks = p.parse();
            let pp = crate::asp::preprocessor::Preprocessor::new();
            let (_inner_config, inner_filtered) = pp.process(&inner_blocks);
            for block in inner_filtered {
                if ctx.response.ended {
                    break;
                }
                handler_clone
                    .handle(block, ctx)
                    .map_err(|e| format!("Execution error in '{}': {}", target, e))?;
            }
            Ok(())
        }));

        // Inject DAP debugger if provided (before block execution)
        context.debugger = debugger;

        // Inject ASP intrinsic objects before block execution
        Self::inject_asp_intrinsic_objects(&mut context);

        // Execute filtered blocks
        let mut response_content = String::new();
        eprintln!("[ASP] block execution starting, count={}", filtered_blocks.len());

        for block in &filtered_blocks {
            eprintln!("[ASP]  -> processing block");
            if context.response.ended {
                break;
            }
            match handler_chain.handle(block, &mut context) {
                Ok(()) => {
                    response_content.push_str(&context.response.buffer);
                }
                Err(e) => {
                    if context.response.ended {
                        break;
                    }
                    response_content.push_str(&context.response.buffer);
                    response_content.push_str(&format!("\n<!-- Error: {} -->\n", e));
                }
            }
            context.flush_response_buffer();
        }

        // Transfer response cookies to headers
        for (name, val) in &context.response.cookies {
            context.response.extra_headers.push((
                "Set-Cookie".to_string(),
                format!("{}={}; path=/", name, val),
            ));
        }

        // Prepend flushed content
        if !context.response.flushed.is_empty() {
            response_content = format!("{}{}", context.response.flushed, response_content);
        }

        // Build response based on context state
        eprintln!("[ASP] block execution done, response_content={} bytes, redirect={:?}, status={:?}",
            response_content.len(), context.response.redirect_url, context.response.status);

        if !context.response.redirect_url.is_empty() {
            response_content.clear();
            Ok(HttpResponse {
                status_line: "302 Found".to_string(),
                content_type: "text/html".to_string(),
                body: String::new(),
                extra_headers: context.response.extra_headers,
            })
        } else {
            Ok(HttpResponse {
                status_line: context.response.status,
                content_type: "text/html".to_string(),
                body: response_content,
                extra_headers: context.response.extra_headers,
            })
        }
    }

    /// Handle a single HTTP connection: read request, process it, and write the response.
    pub async fn handle_connection(
        handler_chain: &Arc<dyn Handler + Send + Sync>,
        stream: &mut tokio::net::TcpStream,
        folder: &str,
        default_document: &str,
        store: &Arc<Store>,
    ) -> Result<(), ASPError> {
        let request = Self::read_request(stream).await?;
        let response = Self::process_request(request, handler_chain, folder, default_document, store, None).await?;
        Self::write_response(stream, &response).await
    }
}
