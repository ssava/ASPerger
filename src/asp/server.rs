use std::fs;
use std::io::{Read, Write};
use std::net::TcpStream;
use super::parser::{AspBlock, AspParser};
use crate::vbscript::{VBScriptInterpreter, ExecutionContext};

pub struct AspServer {
    interpreter: VBScriptInterpreter,
}

impl AspServer {
    pub fn new() -> Self {
        AspServer {
            interpreter: VBScriptInterpreter,
        }
    }

    pub fn start(&self, port: u16) -> std::io::Result<()> {
        let listener = std::net::TcpListener::bind(format!("127.0.0.1:{}", port))?;
        println!("Server in ascolto sulla porta {}", port);

        for stream in listener.incoming() {
            match stream {
                Ok(stream) => {
                    if let Err(e) = self.handle_connection(stream) {
                        eprintln!("Errore nella gestione della connessione: {}", e);
                    }
                }
                Err(e) => eprintln!("Errore di connessione: {}", e),
            }
        }

        Ok(())
    }

    fn handle_connection(&self, mut stream: TcpStream) -> std::io::Result<()> {
        let mut buffer = [0; 1024];
        stream.read(&mut buffer)?;

        let content = fs::read_to_string("test.asp")
            .unwrap_or_else(|_| "<%Response.Write(\"Hello World\")%>".to_string());
        
        let parser = AspParser::new(content);
        let blocks = parser.parse();
        let mut context = ExecutionContext::new();
        let mut response_content = String::new();

        for block in blocks {
            match block {
                AspBlock::Html(html) => response_content.push_str(&html),
                AspBlock::Code(code) => {
                    match self.interpreter.execute(&code, &mut context) {
                        Ok(_) => {
                            response_content.push_str(&context.response_buffer);
                            context.response_buffer.clear();
                        },
                        Err(e) => {
                            eprintln!("Error executing code: {}", e);
                            response_content.push_str(&format!("<!-- Error: {} -->", e));
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

        stream.write(response.as_bytes())?;
        stream.flush()?;

        Ok(())
    }
}