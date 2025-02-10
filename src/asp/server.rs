use crate::asp::asp_error::ASPError;
use crate::asp::config::Config;
use crate::asp::handler::{CodeHandler, Handler, HtmlHandler};
use crate::asp::parser::AspParser;
use crate::vbscript::{ExecutionContext, VBScriptInterpreter};
use std::collections::HashMap;
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
        let listener = TcpListener::bind(format!("127.0.0.1:{}", self.config.port)).await?;
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

    fn parse_request(request: &[u8]) -> Option<(String, String, HashMap<String, String>)> {
        let request_str = String::from_utf8_lossy(request);
        let mut lines = request_str.lines();

        // Parse request line
        let first_line = lines.next()?;
        let mut parts = first_line.split_whitespace();
        let method = parts.next()?.to_string();
        let path = parts
            .next()?
            .trim_start_matches('/')
            .split('?')
            .next()?
            .to_string();

        // Parse headers
        let mut headers = HashMap::new();
        for line in lines {
            if line.is_empty() {
                break;
            }
            if let Some((key, value)) = line.split_once(':') {
                headers.insert(key.trim().to_lowercase(), value.trim().to_string());
            }
        }

        Some((method, path, headers))
    }

    async fn send_response(
        stream: &mut tokio::net::TcpStream, 
        status: u16,
        content_type: &str,
        content: &str
    ) -> Result<(), ASPError> {
        let response = format!(
            "HTTP/1.1 {} {}\r\n\
             Content-Type: {}; charset=utf-8\r\n\
             Content-Length: {}\r\n\
             \r\n\
             {}",
            status,
            if status == 200 { "OK" } else { "Not Found" },
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

    async fn handle_connection(
        handler_chain: &Arc<dyn Handler + Send + Sync>,
        stream: &mut tokio::net::TcpStream,
        folder: &str,
    ) -> Result<(), ASPError> {
        let mut buffer = [0; 4096];
        stream.read(&mut buffer).await.map_err(|e| {
            ASPError::new(500, format!("Error reading from client: {}", e))
        })?;

        let (method, path, headers) = Self::parse_request(&buffer)
            .unwrap_or_else(|| (String::from("GET"), String::from("index.asp"), HashMap::new()));

        // Only handle GET requests for now
        if method != "GET" {
            return Self::send_response(
                stream, 
                405,
                "text/plain",
                "Method not allowed"
            ).await;
        }

        let file_path = format!("{}/{}", folder, if path.is_empty() { "index.asp".to_string() } else { path });
        let path = Path::new(&file_path);
        
        // Ensure the path is within the allowed folder
        let canonical_path = path.canonicalize().map_err(|_| {
            ASPError::new(404, "File not found".to_string())
        })?;
        
        let canonical_folder = Path::new(folder).canonicalize().map_err(|_| {
            ASPError::new(500, "Server configuration error".to_string())
        })?;
        
        if !canonical_path.starts_with(canonical_folder) {
            return Self::send_response(
                stream,
                403,
                "text/plain",
                "Forbidden"
            ).await;
        }

        // Read and process the file
        let content = match std::fs::read_to_string(&file_path) {
            Ok(content) => content,
            Err(_) => {
                return Self::send_response(
                    stream,
                    404,
                    "text/plain",
                    &format!("Page not found: {}", file_path)
                ).await;
            }
        };

        // Handle ASP files
        if file_path.ends_with(".asp") {
            let parser = AspParser::new(content);
            let blocks = parser.parse();

            let mut context = ExecutionContext::new();
            let mut response_content = String::new();

            for block in blocks {
                if let Err(e) = handler_chain.handle(&block, &mut context) {
                    response_content.push_str(&format!("<!-- Error: {} -->", e));
                } else {
                    response_content.push_str(&context.response_buffer);
                    context.flush_response_buffer();
                }
            }

            Self::send_response(stream, 200, "text/html", &response_content).await
        } else {
            // Serve static files
            let content_type = match path.extension().and_then(|e| e.to_str()) {
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
