use std::sync::Arc;
use tokio::net::TcpListener;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use crate::vbscript::{VBScriptInterpreter, ExecutionContext};
use crate::asp::parser::AspParser;
use crate::asp::handler::{Handler, HtmlHandler, CodeHandler};
use crate::asp::asp_error::ASPError;
use crate::asp::config::Config;

pub struct AspServer {
    interpreter: Arc<VBScriptInterpreter>,
    handler_chain: Arc<dyn Handler + Send + Sync>, // Handler chain
    config: Config, // Configurazione del server
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
            interpreter,
            handler_chain: Arc::new(html_handler), // Set the handler chain
            config, // Salva la configurazione
        }
    }

    pub async fn start(&self) -> std::io::Result<()> {
        let listener = TcpListener::bind(format!("127.0.0.1:{}", self.config.port)).await?;
        println!(
            "Server in ascolto sulla porta {} e serve file dalla cartella {}",
            self.config.port, self.config.folder
        );

        loop {
            let (mut stream, _) = listener.accept().await?;
            let handler_chain = Arc::clone(&self.handler_chain); // Clone the Arc
            let folder = self.config.folder.clone(); // Clona la cartella dei file

            tokio::spawn(async move {
                if let Err(e) = Self::handle_connection(&handler_chain, &mut stream, &folder).await {
                    eprintln!("Errore nella gestione della connessione: {}", e.to_string());
                }
            });
        }
    }

    async fn handle_connection(
        handler_chain: &Arc<dyn Handler + Send + Sync>,
        stream: &mut tokio::net::TcpStream,
        folder: &str,
    ) -> Result<(), ASPError> {
        let mut buffer = [0; 1024];
        stream.read(&mut buffer).await.map_err(|e| {
            ASPError::new(500, format!("Errore durante la lettura dal client: {}", e))
        })?;

        // Leggi il contenuto del file ASP dalla cartella specificata
        let file_path = format!("{}/test.asp", folder); // Esempio: ./test.asp
        let content = std::fs::read_to_string(&file_path).unwrap_or_else(|_| {
            eprintln!("File non trovato: {}. Usando contenuto di default.", file_path);
            "<%Response.Write(\"Hello World\")%>".to_string()
        });

        let parser = AspParser::new(content);
        let blocks = parser.parse();

        let mut context = ExecutionContext::new();
        let mut response_content = String::new();

        for block in blocks {
            if let Err(e) = handler_chain.handle(&block, &mut context) {
                // Include the error in the response for debugging
                response_content.push_str(&format!(
                    "<!-- Error: {} -->",
                    e
                ));
            } else {
                response_content.push_str(&context.response_buffer);
                context.flush_response_buffer();
            }
        }

        let response = format!(
            "HTTP/1.1 200 OK\r\n\
             Content-Type: text/html; charset=utf-8\r\n\
             Content-Length: {}\r\n\
             \r\n\
             {}",
            response_content.len(),
            response_content
        );

        stream.write_all(response.as_bytes()).await.map_err(|e| {
            ASPError::new(500, format!("Errore durante la scrittura della risposta: {}", e))
        })?;

        stream.flush().await.map_err(|e| {
            ASPError::new(500, format!("Errore durante lo svuotamento del buffer: {}", e))
        })?;

        Ok(())
    }
}