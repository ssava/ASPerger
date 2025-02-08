use std::sync::Arc;
use tokio::net::TcpListener;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use crate::vbscript::vbs_error::VBSError;
use crate::vbscript::{VBScriptInterpreter, ExecutionContext};
use crate::asp::parser::AspParser;
use crate::asp::parser::AspBlock;

pub struct AspServer {
    interpreter: Arc<VBScriptInterpreter>,
}

impl AspServer {
    pub fn new() -> Self {
        AspServer {
            interpreter: Arc::new(VBScriptInterpreter),
        }
    }

    pub async fn start(&self, port: u16) -> std::io::Result<()> {
        let listener = TcpListener::bind(format!("127.0.0.1:{}", port)).await?;
        println!("Server in ascolto sulla porta {}", port);
        loop {
            let (mut stream, _) = listener.accept().await?;
            let interpreter = Arc::clone(&self.interpreter);
            tokio::spawn(async move {
                if let Err(e) = Self::handle_connection(&interpreter, &mut stream).await {
                    eprintln!("Errore nella gestione della connessione: {}", e.to_string());
                }
            });
        }
    }

    async fn handle_connection(
        interpreter: &VBScriptInterpreter,
        stream: &mut tokio::net::TcpStream,
    ) -> Result<(), VBSError> { // Cambia il tipo di ritorno in Result<(), VBSError>
        let mut buffer = [0; 1024];
        if let Err(e) = stream.read(&mut buffer).await.map_err(|e| VBSError::new(500, format!("Errore durante la lettura dal client: {}", e))) {
            return Err(e);
        }

        // Leggi il contenuto del file ASP o usa un fallback
        let content = std::fs::read_to_string("test.asp")
            .unwrap_or_else(|_| "<%Response.Write(\"Hello World\")%>".to_string());

        let parser = AspParser::new(content);
        let blocks = parser.parse();

        let mut context = ExecutionContext::new();
        let mut response_content = String::new();

        for block in blocks {
            match block {
                AspBlock::Html(html) => response_content.push_str(&html),
                AspBlock::Code(code) => {
                    match interpreter.execute(&code, &mut context) {
                        Ok(_) => {
                            response_content.push_str(&context.response_buffer);
                            context.flush_response_buffer();
                        }
                        Err(e) => {
                            eprintln!("Error executing code: {}", e.to_string());
                            response_content.push_str(&format!(
                                "<!-- Error: [Codice {}]: {} -->",
                                e.code, e.message
                            ));
                        }
                    }
                }
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

        if let Err(e) = stream.write_all(response.as_bytes()).await.map_err(|e| VBSError::new(500, format!("Errore durante la scrittura della risposta: {}", e))) {
            return Err(e);
        }

        if let Err(e) = stream.flush().await.map_err(|e| VBSError::new(500, format!("Errore durante lo svuotamento del buffer: {}", e))) {
            return Err(e);
        }

        Ok(())
    }
}