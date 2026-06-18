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

use axum::{
    body::Body,
    extract::{Extension, Request},
    http::StatusCode,
    response::Response,
    routing::any,
    Router,
};

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
    /// Response body bytes.
    pub body: Vec<u8>,
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
        let folder = if asp_cfg.folder.is_empty() || asp_cfg.folder.trim_end_matches('/').is_empty() {
            self.config.folder.trim_end_matches('/').to_string()
        } else {
            asp_cfg.folder.trim_end_matches('/').to_string()
        };
        let default_doc = asp_cfg.default_document.clone();
        let dir_listing = asp_cfg.directory_listing;

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
                    Self::handle_connection(&handler_chain, &mut stream, &folder, &default_doc, &store, dir_listing).await
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

    /// Minimal percent-encoding for filenames in directory listing links.
    fn url_encode(s: &str) -> String {
        let mut result = String::with_capacity(s.len());
        for b in s.bytes() {
            match b {
                b'-' | b'_' | b'.' | b'~' | b'/' => result.push(b as char),
                b if b.is_ascii_alphanumeric() => result.push(b as char),
                b' ' => result.push_str("%20"),
                _ => result.push_str(&format!("%{:02X}", b)),
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
        let header = format!(
            "HTTP/1.1 {}\r\nContent-Length: {}\r\n",
            response.status_line,
            response.body.len(),
        );
        let mut buf = Vec::new();
        buf.extend_from_slice(header.as_bytes());
        if !has_content_type {
            buf.extend_from_slice(
                format!(
                    "Content-Type: {}; charset=utf-8\r\n",
                    response.content_type
                )
                .as_bytes(),
            );
        }
        for (key, value) in &response.extra_headers {
            buf.extend_from_slice(format!("{}: {}\r\n", key, value).as_bytes());
        }
        buf.extend_from_slice(b"\r\n");
        buf.extend_from_slice(&response.body);

        stream
            .write_all(&buf)
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
        directory_listing: bool,
    ) -> Result<HttpResponse, ASPError> {
        // Build raw path — no default document substitution yet
        let raw_path = format!("{}/{}", folder, request.path);

        let canonical_path = match Path::new(&raw_path).canonicalize() {
            Ok(p) => p,
            Err(e) => {
                let err = ASPError::new(404, &format!("File not found: {} (folder={}, path={}, error={})", raw_path, folder, request.path, e));
                return Ok(HttpResponse {
                    status_line: "404 Not Found".to_string(),
                    content_type: "text/html".to_string(),
                    body: err.render_html().into_bytes(),
                    extra_headers: Vec::new(),
                });
            }
        };

        let canonical_folder = Path::new(folder)
            .canonicalize()
            .map_err(|_| ASPError::new(500, "Server configuration error".to_string()))?;

        if !canonical_path.starts_with(&canonical_folder) {
            let err = ASPError::new(403, "Forbidden: access denied");
            return Ok(HttpResponse {
                status_line: "403 Forbidden".to_string(),
                content_type: "text/html".to_string(),
                body: err.render_html().into_bytes(),
                extra_headers: Vec::new(),
            });
        }

        // Determine the actual file path:
        // - If it's a directory, try default document, then directory listing, else 404
        // - If it's a file, use raw_path directly
        let file_path = if canonical_path.is_dir() {
            let default_path = canonical_path.join(default_document);
            if default_path.is_file() {
                default_path.to_string_lossy().to_string()
            } else if directory_listing {
                return Ok(Self::generate_directory_listing(&canonical_path, &canonical_folder, &request.path));
            } else {
            return Ok(HttpResponse {
                status_line: "404 Not Found".to_string(),
                content_type: "text/html".to_string(),
                body: format!(
                    "<html><head><title>404 Not Found</title>\
                     <style>body{{font-family:monospace;background:#f8f8f8;padding:2em}}\
                     h1{{color:#c00;font-size:1.5em}}\
                     .code{{color:#666}}\
                     .msg{{background:#fff;border:1px solid #ddd;padding:1em;margin:1em 0}}\
                     </style></head><body>\
                     <h1>404 Not Found</h1>\
                     <p class=\"code\">{}</p>\
                     <div class=\"msg\">{}</div>\
                     </body></html>",
                    Self::html_escape(&request.path),
                    Self::html_escape(&format!(
                        "No default document ({}) found",
                        default_document
                    ))
                ).into_bytes(),
                extra_headers: Vec::new(),
            });
            }
        } else {
            raw_path
        };
        let path_obj = Path::new(&file_path);

        let content = match std::fs::read(&file_path) {
            Ok(content) => content,
            Err(_) => {
                return Ok(HttpResponse {
                    status_line: "404 Not Found".to_string(),
                    content_type: "text/plain".to_string(),
                    body: format!("Page not found: {}", file_path).into_bytes(),
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

        let content = String::from_utf8(content)
            .map_err(|e| ASPError::new(500, format!("Non-UTF8 content in ASP file: {}", e)))?;

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
                    body: ASPError::new(500, e).render_html().into_bytes(),
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
        // Populate standard CGI/ServerVariables
        context.request.headers.insert(
            "script_name".to_string(),
            format!("/{}", request.path),
        );
        context.request.headers.insert(
            "server_name".to_string(),
            request
                .headers
                .get("host")
                .and_then(|h| h.split(':').next())
                .unwrap_or("127.0.0.1")
                .to_string(),
        );
        context
            .request
            .headers
            .insert("request_method".to_string(), request.method.clone());
        context
            .request
            .headers
            .insert("query_string".to_string(), request.query_string.clone());
        context.request.headers.insert(
            "server_port".to_string(),
            request
                .headers
                .get("host")
                .and_then(|h| h.split(':').nth(1))
                .unwrap_or("8080")
                .to_string(),
        );
        context
            .request
            .headers
            .insert("server_protocol".to_string(), "HTTP/1.1".to_string());
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

        for block in &filtered_blocks {
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
        if !context.response.redirect_url.is_empty() {
            response_content.clear();
            Ok(HttpResponse {
                status_line: "302 Found".to_string(),
                content_type: "text/html".to_string(),
                body: Vec::new(),
                extra_headers: context.response.extra_headers,
            })
        } else {
            Ok(HttpResponse {
                status_line: context.response.status,
                content_type: "text/html".to_string(),
                body: response_content.into_bytes(),
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
        directory_listing: bool,
    ) -> Result<(), ASPError> {
        let request = Self::read_request(stream).await?;
        let response = Self::process_request(request, handler_chain, folder, default_document, store, None, directory_listing).await?;
        Self::write_response(stream, &response).await
    }

    /// Start the HTTP server using axum (production path).
    /// Supports graceful shutdown on Ctrl+C.
    pub async fn start_axum(&self, asp_cfg: &AspServerConfig) -> std::io::Result<()> {
        let host = if asp_cfg.host.is_empty() { &self.config.host } else { &asp_cfg.host };
        let port = if asp_cfg.port == 0 { self.config.port } else { asp_cfg.port };
        let folder = if asp_cfg.folder.is_empty() || asp_cfg.folder.trim_end_matches('/').is_empty() {
            self.config.folder.trim_end_matches('/').to_string()
        } else {
            asp_cfg.folder.trim_end_matches('/').to_string()
        };
        let default_doc = asp_cfg.default_document.clone();

        let state = Arc::new(AxumServerState {
            handler_chain: Arc::clone(&self.handler_chain),
            store: Arc::clone(&self.store),
            folder: folder.clone(),
            default_document: default_doc.clone(),
            directory_listing: asp_cfg.directory_listing,
        });

        let app = Router::new()
            .fallback(any(axum_handler))
            .layer(Extension(state));

        let addr: std::net::SocketAddr = format!("{}:{}", host, port)
            .parse()
            .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidInput, e))?;
        println!(
            "Server listening on {} serving files from {} (default document: {})",
            addr, folder, default_doc
        );

        let listener = tokio::net::TcpListener::bind(addr).await?;
        axum::serve(listener, app)
            .with_graceful_shutdown(async {
                tokio::signal::ctrl_c().await.ok();
            })
            .await
            .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e))
    }

    /// Generate an HTML directory listing for the given canonical directory path.
    fn generate_directory_listing(
        dir_path: &std::path::Path,
        root_path: &std::path::Path,
        request_path: &str,
    ) -> HttpResponse {
        let mut rows = String::new();

        // Parent directory link (only if within served root)
        if dir_path != root_path {
            let parent = if request_path.is_empty() || request_path == "/" {
                "".to_string()
            } else {
                let trimmed = request_path.trim_end_matches('/');
                if let Some(pos) = trimmed.rfind('/') {
                    format!("{}/", &trimmed[..=pos])
                } else {
                    "".to_string()
                }
            };
            rows.push_str(&format!(
                "<tr><td><a href=\"{}\">../</a></td><td></td><td></td></tr>\n",
                Self::html_escape(&parent)
            ));
        }

        let read_dir = match std::fs::read_dir(dir_path) {
            Ok(r) => r,
            Err(_) => {
                return HttpResponse {
                    status_line: "500 Internal Server Error".to_string(),
                    content_type: "text/html".to_string(),
                    body: "Unable to read directory".to_string().into_bytes(),
                    extra_headers: vec![(
                        "Content-Security-Policy".to_string(),
                        "default-src 'self'".to_string(),
                    )],
                };
            }
        };

        struct ListEntry {
            name: String,
            is_dir: bool,
            size: u64,
            modified: String,
        }

        let mut entries: Vec<ListEntry> = Vec::new();

        for entry in read_dir.flatten() {
            let name = entry.file_name();
            let name_str = name.to_string_lossy().to_string();

            // Security: skip hidden files and config
            if name_str.starts_with('.') || name_str.eq_ignore_ascii_case("asp.ini") {
                continue;
            }

            let meta = match std::fs::symlink_metadata(entry.path()) {
                Ok(m) => m,
                Err(_) => continue,
            };

            // Skip symlinks — they could point outside the served tree
            if meta.file_type().is_symlink() {
                continue;
            }

            let is_dir = meta.is_dir();
            let size = meta.len();
            let modified = meta
                .modified()
                .ok()
                .map(|t| {
                    let duration = t
                        .duration_since(std::time::UNIX_EPOCH)
                        .unwrap_or_default();
                    // Format as YYYY-MM-DD HH:MM:SS
                    let secs = duration.as_secs();
                    let days = secs / 86400;
                    let time_secs = secs % 86400;
                    let hours = time_secs / 3600;
                    let minutes = (time_secs % 3600) / 60;
                    let seconds = time_secs % 60;
                    // Approximate year (1970 + days/365.25)
                    let year = 1970 + (days as f64 / 365.25) as u64;
                    let month = 1 + ((days % 365) / 30).min(11);
                    let day = 1 + (days % 30).min(30);
                    format!(
                        "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                        year, month, day, hours, minutes, seconds
                    )
                })
                .unwrap_or_default();

            entries.push(ListEntry {
                name: name_str,
                is_dir,
                size,
                modified,
            });
        }

        // Sort: directories first, then files; alphabetically within each group
        entries.sort_by(|a, b| {
            b.is_dir
                .cmp(&a.is_dir)
                .then_with(|| a.name.to_lowercase().cmp(&b.name.to_lowercase()))
        });

        for e in &entries {
            let href = if e.is_dir {
                format!("{}/", Self::url_encode(&e.name))
            } else {
                Self::url_encode(&e.name)
            };
            let size_str = if e.is_dir {
                String::new()
            } else if e.size < 1024 {
                format!("{} B", e.size)
            } else if e.size < 1024 * 1024 {
                format!("{:.1} KB", e.size as f64 / 1024.0)
            } else {
                format!("{:.1} MB", e.size as f64 / (1024.0 * 1024.0))
            };
            rows.push_str(&format!(
                "<tr><td><a href=\"{}\">{}</a></td><td>{}</td><td>{}</td></tr>\n",
                Self::html_escape(&href),
                Self::html_escape(&e.name),
                Self::html_escape(&size_str),
                Self::html_escape(&e.modified),
            ));
        }

        let path_display = format!("/{}", request_path);
        let body = format!(
            "<!DOCTYPE html>\n\
             <html>\n<head>\n\
             <meta charset=\"UTF-8\">\n\
             <title>Index of {}</title>\n\
             <style>\n\
             body{{font-family:monospace;background:#f8f8f8;padding:2em}}\n\
             h1{{color:#c00;font-size:1.5em}}\n\
             .code{{color:#666}}\n\
             .msg{{background:#fff;border:1px solid #ddd;padding:1em;margin:1em 0}}\n\
             table{{width:100%;border-collapse:collapse}}\n\
             th,td{{text-align:left;padding:4px 8px;border-bottom:1px solid #ddd}}\n\
             th{{background:#f5f5f5}}\n\
             tr:hover{{background:#eee}}\n\
             a{{color:#06c;text-decoration:none}}\n\
             a:hover{{text-decoration:underline}}\n\
             </style>\n\
             </head>\n<body>\n\
             <h1>Index of {}</h1>\n\
             <div class=\"msg\">\n\
             <table>\n\
             <tr><th>Name</th><th>Size</th><th>Last Modified</th></tr>\n\
             {} \
             </table>\n\
             </div>\n</body>\n</html>",
            Self::html_escape(&path_display),
            Self::html_escape(&path_display),
            rows,
        );

        HttpResponse {
            status_line: "200 OK".to_string(),
            content_type: "text/html".to_string(),
            body: body.into_bytes(),
            extra_headers: vec![(
                "Content-Security-Policy".to_string(),
                "default-src 'self'".to_string(),
            )],
        }
    }

    /// Minimal HTML-entity escaping for text content.
    fn html_escape(s: &str) -> String {
        let mut out = String::with_capacity(s.len());
        for c in s.chars() {
            match c {
                '&' => out.push_str("&amp;"),
                '<' => out.push_str("&lt;"),
                '>' => out.push_str("&gt;"),
                '"' => out.push_str("&quot;"),
                '\'' => out.push_str("&#x27;"),
                _ => out.push(c),
            }
        }
        out
    }
}

/// Shared state injected into all axum handlers.
struct AxumServerState {
    handler_chain: Arc<dyn Handler + Send + Sync>,
    store: Arc<Store>,
    folder: String,
    default_document: String,
    directory_listing: bool,
}

/// Axum request handler: converts axum Request → internal HttpRequest,
/// calls process_request, then converts internal HttpResponse → axum Response.
async fn axum_handler(
    Extension(state): Extension<Arc<AxumServerState>>,
    req: Request,
) -> Response<Body> {
    let (parts, body) = req.into_parts();
    let body_bytes = axum::body::to_bytes(body, 10 * 1024 * 1024)
        .await
        .unwrap_or_default();

    let headers: AHashMap<String, String> = parts
        .headers
        .iter()
        .map(|(k, v)| {
            (
                k.as_str().to_lowercase(),
                v.to_str().unwrap_or("").to_string(),
            )
        })
        .collect();

    let http_request = crate::asp::server::HttpRequest {
        method: parts.method.to_string(),
        path: parts
            .uri
            .path()
            .trim_start_matches('/')
            .to_string(),
        query_string: parts.uri.query().unwrap_or("").to_string(),
        cookies: headers
            .get("cookie")
            .map(|c| AspServer::parse_cookies(c))
            .unwrap_or_default(),
        headers,
        body: body_bytes.to_vec(),
    };

    match AspServer::process_request(
        http_request,
        &state.handler_chain,
        &state.folder,
        &state.default_document,
        &state.store,
        None,
        state.directory_listing,
    )
    .await
    {
        Ok(http_resp) => convert_response(http_resp),
        Err(e) => Response::builder()
            .status(StatusCode::INTERNAL_SERVER_ERROR)
            .body(Body::from(format!("Internal error: {}", e)))
            .unwrap(),
    }
}

/// Convert our internal HttpResponse to an axum Response<Body>.
fn convert_response(resp: crate::asp::server::HttpResponse) -> Response<Body> {
    let status_code = parse_status_code(&resp.status_line);
    let has_content_type = resp
        .extra_headers
        .iter()
        .any(|(k, _)| k.eq_ignore_ascii_case("content-type"));

    let mut builder = Response::builder()
        .status(status_code)
        .header("Content-Length", resp.body.len().to_string());
    if !has_content_type {
        builder = builder.header(
            "Content-Type",
            format!("{}; charset=utf-8", resp.content_type),
        );
    }
    for (key, value) in &resp.extra_headers {
        builder = builder.header(key.as_str(), value.as_str());
    }
    builder
        .body(Body::from(resp.body))
        .unwrap_or_else(|_| {
            Response::builder()
                .status(StatusCode::INTERNAL_SERVER_ERROR)
                .body(Body::from("Response build error"))
                .unwrap()
        })
}

/// Parse a status line string like "200 OK" into a StatusCode.
fn parse_status_code(status_line: &str) -> StatusCode {
    let code = status_line
        .split_whitespace()
        .next()
        .and_then(|s| s.parse::<u16>().ok())
        .unwrap_or(500);
    StatusCode::from_u16(code).unwrap_or(StatusCode::INTERNAL_SERVER_ERROR)
}
