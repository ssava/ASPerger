use std::fs;
use std::io::{Read, Write};
use std::net::{TcpListener, TcpStream};
use regex::Regex;
use std::collections::HashMap;

#[derive(Clone, Debug)]
enum VBValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Null,
}

impl ToString for VBValue {
    fn to_string(&self) -> String {
        match self {
            VBValue::String(s) => s.clone(),
            VBValue::Number(n) => n.to_string(),
            VBValue::Boolean(b) => b.to_string(),
            VBValue::Null => "null".to_string(),
        }
    }
}

#[derive(Default)]
struct ExecutionContext {
    variables: HashMap<String, VBValue>,
    response_buffer: String,
}

impl ExecutionContext {
    fn new() -> Self {
        ExecutionContext {
            variables: HashMap::new(),
            response_buffer: String::new(),
        }
    }

    fn write(&mut self, content: &str) {
        self.response_buffer.push_str(content);
    }

    fn set_variable(&mut self, name: &str, value: VBValue) {
        self.variables.insert(name.to_string(), value);
    }

    fn get_variable(&self, name: &str) -> Option<VBValue> {
        self.variables.get(name).cloned()
    }
}

struct VBScriptInterpreter;

impl VBScriptInterpreter {
    fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let code = code.trim();

        // Handle each line separately
        for line in code.split('\n') {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }

            // Skip comments
            if line.starts_with('\'') || line.to_lowercase().starts_with("rem") {
                continue;
            }

            if line.to_lowercase().starts_with("response.write") {
                self.handle_response_write(line, context)?;
            } else if line.to_lowercase().starts_with("dim") {
                self.handle_dim(line, context)?;
            } else if line.contains('=') {
                self.handle_assignment(line, context)?;
            } else if line.to_lowercase().starts_with("if") {
                self.handle_if(line, context)?;
            }
        }
        Ok(())
    }

    fn handle_response_write(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let content = code
            .trim()
            .strip_prefix("Response.Write")
            .unwrap_or("")
            .trim();

        // Remove outer parentheses if present
        let content = if content.starts_with('(') && content.ends_with(')') {
            &content[1..content.len()-1]
        } else {
            content
        };

        // Trim again after removing parentheses
        let content = content.trim();

        // Handle string literal with quotes
        if content.starts_with('"') && content.ends_with('"') {
            let content = &content[1..content.len()-1];
            context.write(content);
            return Ok(());
        }

        // Handle variable
        if let Some(value) = context.get_variable(content) {
            context.write(&value.to_string());
            return Ok(());
        }

        // Handle number
        if let Ok(num) = content.parse::<f64>() {
            context.write(&num.to_string());
            return Ok(());
        }

        Err(format!("Valore non valido per Response.Write: {}", content))
    }

    fn handle_dim(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let var_names = code
            .trim()
            .strip_prefix("Dim")
            .unwrap_or("")
            .split(',')
            .map(|s| s.trim())
            .filter(|s| !s.is_empty());

        for var_name in var_names {
            context.set_variable(var_name, VBValue::Null);
        }
        
        Ok(())
    }

    fn handle_assignment(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let parts: Vec<&str> = code.split('=').collect();
        if parts.len() != 2 {
            return Err("Assegnazione non valida".to_string());
        }

        let var_name = parts[0].trim();
        let value = parts[1].trim();

        // Gestione valore stringa
        if value.starts_with('"') && value.ends_with('"') {
            let string_value = value.trim_matches('"').to_string();
            context.set_variable(var_name, VBValue::String(string_value));
            return Ok(());
        }

        // Gestione valore numerico
        if let Ok(num) = value.parse::<f64>() {
            context.set_variable(var_name, VBValue::Number(num));
            return Ok(());
        }

        // Gestione booleani
        match value.to_lowercase().as_str() {
            "true" => {
                context.set_variable(var_name, VBValue::Boolean(true));
                return Ok(());
            }
            "false" => {
                context.set_variable(var_name, VBValue::Boolean(false));
                return Ok(());
            }
            _ => {}
        }

        // Gestione riferimento ad altra variabile
        if let Some(var_value) = context.get_variable(value) {
            context.set_variable(var_name, var_value);
            return Ok(());
        }

        Err(format!("Valore non valido per l'assegnazione: {}", value))
    }

    fn handle_if(&self, code: &str, context: &mut ExecutionContext) -> Result<(), String> {
        let if_pattern = Regex::new(r"If\s+(.+?)\s+Then\s+(.+?)\s+End If").unwrap();
        
        if let Some(caps) = if_pattern.captures(code) {
            let condition = caps.get(1).unwrap().as_str();
            let then_code = caps.get(2).unwrap().as_str();

            if self.evaluate_condition(condition, context)? {
                self.execute(then_code, context)?;
            }
            return Ok(());
        }

        Err("Sintassi If non valida".to_string())
    }

    fn evaluate_condition(&self, condition: &str, context: &mut ExecutionContext) -> Result<bool, String> {
        let parts: Vec<&str> = condition.split_whitespace().collect();
        if parts.len() != 3 {
            return Err("Condizione non valida".to_string());
        }

        let left = if let Some(value) = context.get_variable(parts[0]) {
            value
        } else if let Ok(num) = parts[0].parse::<f64>() {
            VBValue::Number(num)
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        let operator = parts[1];
        let right = if parts[2].starts_with('"') {
            VBValue::String(parts[2].trim_matches('"').to_string())
        } else if let Ok(num) = parts[2].parse::<f64>() {
            VBValue::Number(num)
        } else if let Some(value) = context.get_variable(parts[2]) {
            value
        } else {
            return Err("Variabile o valore non trovato".to_string());
        };

        match operator {
            "=" => Ok(self.compare_values(&left, &right)),
            ">" => Ok(self.compare_greater(&left, &right)),
            "<" => Ok(self.compare_less(&left, &right)),
            _ => Err("Operatore non supportato".to_string()),
        }
    }

    fn compare_values(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::String(a), VBValue::String(b)) => a == b,
            (VBValue::Number(a), VBValue::Number(b)) => (a - b).abs() < f64::EPSILON,
            (VBValue::Boolean(a), VBValue::Boolean(b)) => a == b,
            _ => false,
        }
    }

    fn compare_greater(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::Number(a), VBValue::Number(b)) => a > b,
            _ => false,
        }
    }

    fn compare_less(&self, left: &VBValue, right: &VBValue) -> bool {
        match (left, right) {
            (VBValue::Number(a), VBValue::Number(b)) => a < b,
            _ => false,
        }
    }
}

#[derive(Debug)]
enum AspBlock {
    Html(String),
    Code(String),
}

struct AspParser {
    content: String,
}

impl AspParser {
    fn new(content: String) -> Self {
        AspParser { content }
    }

    fn parse(&self) -> Vec<AspBlock> {
        let mut blocks = Vec::new();
        let re = Regex::new(r"<%(?s)(.*?)%>").unwrap();
        let mut last_end = 0;

        for cap in re.captures_iter(&self.content) {
            let whole_match = cap.get(0).unwrap();
            let code = cap.get(1).map_or("", |m| m.as_str());
            
            if whole_match.start() > last_end {
                let html = &self.content[last_end..whole_match.start()];
                if !html.trim().is_empty() {
                    blocks.push(AspBlock::Html(html.to_string()));
                }
            }

            if !code.trim().is_empty() {
                blocks.push(AspBlock::Code(code.trim().to_string()));
            }
            
            last_end = whole_match.end();
        }

        if last_end < self.content.len() {
            let html = &self.content[last_end..];
            if !html.trim().is_empty() {
                blocks.push(AspBlock::Html(html.to_string()));
            }
        }

        blocks
    }
}

struct AspServer {
    interpreter: VBScriptInterpreter,
}

impl AspServer {
    fn new() -> Self {
        AspServer {
            interpreter: VBScriptInterpreter,
        }
    }

    fn start(&self, port: u16) -> std::io::Result<()> {
        let listener = TcpListener::bind(format!("127.0.0.1:{}", port))?;
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

fn main() {
    let server = AspServer::new();
    if let Err(e) = server.start(8080) {
        eprintln!("Errore del server: {}", e);
    }
}