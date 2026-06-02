use ahash::AHashMap;
use crate::asp::asp_error::ASPError;
use crate::asp::config::Config;
use crate::asp::handler::{CodeHandler, Handler, HtmlHandler};
use crate::asp::parser::AspParser;
use crate::vbscript::{ExecutionContext, VBScriptInterpreter};
use std::path::Path;
use std::sync::Arc;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::net::TcpListener;

pub struct AspServer {
    handler_chain: Arc<dyn Handler + Send + Sync>, // Handler chain
    config: Config,                                // Configurazione del server
}

impl AspServer {
    pub fn new(config: Config) -> Self {
        let interpreter = Arc::new(VBScriptInterpreter);

        // Create handlers
        let mut html_handler = HtmlHandler::new();
        let code_handler = CodeHandler::new(Arc::clone(&interpreter));

        // Build the chain
        html_handler.set_next(Arc::new(code_handler));

        AspServer {
            handler_chain: Arc::new(html_handler), // Set the handler chain
            config,                                // Salva la configurazione
        }
    }

    pub async fn start(&self) -> std::io::Result<()> {
        let listener = TcpListener::bind(format!("{}:{}", self.config.host, self.config.port)).await?;
        println!(
            "Server listening on port {} serving files from {}",
            self.config.port, self.config.folder
        );

        loop {
            let (mut stream, _) = listener.accept().await?;
            let handler_chain = Arc::clone(&self.handler_chain);
            let folder = self.config.folder.clone();

            tokio::spawn(async move {
                if let Err(e) = Self::handle_connection(&handler_chain, &mut stream, &folder).await {
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
                    let hi = chars.next().and_then(|c| (c as char).to_digit(16)).unwrap_or(0);
                    let lo = chars.next().and_then(|c| (c as char).to_digit(16)).unwrap_or(0);
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
            let d = SystemTime::now().duration_since(UNIX_EPOCH).unwrap_or_default();
            (d.as_nanos() & 0xFFFFFFFF) as u64
        };
        format!("{:04x}", val)
    }

    async fn send_response(
        stream: &mut tokio::net::TcpStream, 
        status: u16,
        content_type: &str,
        content: &str
    ) -> Result<(), ASPError> {
        let reason = match status {
            200 => "OK",
            302 => "Found",
            400 => "Bad Request",
            403 => "Forbidden",
            404 => "Not Found",
            405 => "Method Not Allowed",
            500 => "Internal Server Error",
            _ => "Unknown",
        };
        let response = format!(
            "HTTP/1.1 {} {}\r\n\
             Content-Type: {}; charset=utf-8\r\n\
             Content-Length: {}\r\n\
             \r\n\
             {}",
            status,
            reason,
            content_type,
            content.len(),
            content
        );

        stream.write_all(response.as_bytes()).await.map_err(|e| {
            ASPError::new(500, format!("Error writing response: {}", e))
        })?;

        stream.flush().await.map_err(|e| {
            ASPError::new(500, format!("Error flushing buffer: {}", e))
        })?;

        Ok(())
    }

    async fn send_response_with_headers(
        stream: &mut tokio::net::TcpStream,
        status_line: &str,
        content_type: &str,
        content: &str,
        extra_headers: &[(String, String)],
    ) -> Result<(), ASPError> {
        let has_content_type = extra_headers.iter().any(|(k, _)| k.eq_ignore_ascii_case("content-type"));
        let mut response = format!(
            "HTTP/1.1 {}\r\n\
             Content-Length: {}\r\n",
            status_line,
            content.len(),
        );
        if !has_content_type {
            response.push_str(&format!("Content-Type: {}; charset=utf-8\r\n", content_type));
        }
        for (key, value) in extra_headers {
            response.push_str(&format!("{}: {}\r\n", key, value));
        }
        response.push_str("\r\n");
        response.push_str(content);

        stream.write_all(response.as_bytes()).await.map_err(|e| {
            ASPError::new(500, format!("Error writing response: {}", e))
        })?;

        stream.flush().await.map_err(|e| {
            ASPError::new(500, format!("Error flushing buffer: {}", e))
        })?;

        Ok(())
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

    async fn handle_connection(
        handler_chain: &Arc<dyn Handler + Send + Sync>,
        stream: &mut tokio::net::TcpStream,
        folder: &str,
    ) -> Result<(), ASPError> {
        use tokio::io::AsyncBufReadExt;
        use tokio::io::BufReader;

        let mut reader = BufReader::new(&mut *stream);
        let mut request_buf = Vec::new();
        let mut headers = AHashMap::new();

        // Read request line
        let mut request_line = String::new();
        reader.read_line(&mut request_line).await.map_err(|e| {
            ASPError::new(500, format!("Error reading request line: {}", e))
        })?;
        request_buf.extend_from_slice(request_line.as_bytes());
        let mut parts = request_line.split_whitespace();
        let method = parts.next().unwrap_or("GET").to_string();
        let full_path = parts.next().unwrap_or("/").to_string();
        let (path, query_string) = match full_path.split_once('?') {
            Some((p, q)) => (p.trim_start_matches('/').to_string(), q.to_string()),
            None => (full_path.trim_start_matches('/').to_string(), String::new()),
        };

        // Read headers line by line
        loop {
            let mut line = String::new();
            let n = reader.read_line(&mut line).await.map_err(|e| {
                ASPError::new(500, format!("Error reading header: {}", e))
            })?;
            if n <= 1 || line.trim().is_empty() {
                break;
            }
            request_buf.extend_from_slice(line.as_bytes());
            if let Some((key, value)) = line.split_once(':') {
                headers.insert(key.trim().to_lowercase(), value.trim().to_string());
            }
        }

        // Parse cookies from headers
        let cookies = headers
            .get("cookie")
            .map(|c| Self::parse_cookies(c))
            .unwrap_or_default();

        // Read body based on Content-Length
        let content_length: usize = headers
            .get("content-length")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        let body = if content_length > 0 {
            let mut body = vec![0u8; content_length];
            reader.read_exact(&mut body).await.map_err(|e| {
                ASPError::new(500, format!("Error reading body: {}", e))
            })?;
            body
        } else {
            Vec::new()
        };

        let params = Self::parse_query_string(&query_string);

        // Support basic file lookup
        let file_path = format!(
            "{}/{}",
            folder,
            if path.is_empty() {
                "index.asp".to_string()
            } else {
                path.clone()
            }
        );
        let path_obj = Path::new(&file_path);

        // Ensure the path is within the allowed folder
        let canonical_path = path_obj.canonicalize().map_err(|_| {
            ASPError::new(404, "File not found".to_string())
        })?;

        let canonical_folder = Path::new(folder).canonicalize().map_err(|_| {
            ASPError::new(500, "Server configuration error".to_string())
        })?;

        if !canonical_path.starts_with(canonical_folder) {
            return Self::send_response(stream, 403, "text/plain", "Forbidden").await;
        }

        // Read and process the file
        let content = match std::fs::read_to_string(&file_path) {
            Ok(content) => content,
            Err(_) => {
                return Self::send_response(
                    stream,
                    404,
                    "text/plain",
                    &format!("Page not found: {}", file_path),
                )
                .await;
            }
        };

        // Handle ASP files
        if file_path.ends_with(".asp") {
            let parser = AspParser::new(content);
            let blocks = parser.parse();

            let mut context = ExecutionContext::new();

            // Populate request data
            context.request_method = method.clone();
            context.request_path = path.clone();
            context.request_query_string = query_string.clone();
            context.request_params = params;
            context.request_headers = headers.clone();
            context.request_cookies = cookies;

            // Parse form data for POST
            let content_type = headers.get("content-type").cloned().unwrap_or_default();
            if method.eq_ignore_ascii_case("POST")
                && content_type.contains("application/x-www-form-urlencoded")
            {
                context.request_form = Self::parse_form_body(&body);
            }

            // Session handling
            let existing_session = context
                .request_cookies
                .get("ASPSESSIONID")
                .cloned()
                .unwrap_or_default();
            let session_was_new = existing_session.is_empty();
            if session_was_new {
                context.session_id = Self::generate_session_id();
            } else {
                context.session_id = existing_session;
            }

            // Build session cookie for response (only for new sessions)
            if session_was_new && !context.session_id.is_empty() {
                context.response_extra_headers.push((
                    "Set-Cookie".to_string(),
                    format!("ASPSESSIONID={}; path=/", context.session_id),
                ));
            }

            let mut response_content = String::new();

            for block in blocks {
                if context.response_ended {
                    break;
                }
                match handler_chain.handle(&block, &mut context) {
                    Ok(()) => {
                        response_content.push_str(&context.response_buffer);
                    }
                    Err(e) => {
                        if context.response_ended {
                            break;
                        }
                        response_content.push_str(&context.response_buffer);
                        response_content.push_str(&format!("\n<!-- Error: {} -->\n", e));
                    }
                }
                context.flush_response_buffer();
            }

            // Transfer response cookies to headers
            for (name, val) in &context.response_cookies {
                context.response_extra_headers.push((
                    "Set-Cookie".to_string(),
                    format!("{}={}; path=/", name, val),
                ));
            }

            // Prepend flushed content
            if !context.response_flushed.is_empty() {
                response_content = format!("{}{}", context.response_flushed, response_content);
            }

            // Build response based on context state
            if !context.response_redirect_url.is_empty() {
                response_content.clear();
                Self::send_response_with_headers(
                    stream,
                    "302 Found",
                    "text/html",
                    "",
                    &context.response_extra_headers,
                )
                .await
            } else {
                Self::send_response_with_headers(
                    stream,
                    &context.response_status,
                    "text/html",
                    &response_content,
                    &context.response_extra_headers,
                )
                .await
            }
        } else {
            // Serve static files
            let content_type = match path_obj.extension().and_then(|e| e.to_str()) {
                Some("html") | Some("htm") => "text/html",
                Some("css") => "text/css",
                Some("js") => "application/javascript",
                Some("txt") => "text/plain",
                _ => "application/octet-stream",
            };

            Self::send_response(stream, 200, content_type, &content).await
        }
    }
}
