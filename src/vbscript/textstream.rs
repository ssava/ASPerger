use std::fs::File;
use std::io::{BufRead, BufReader, BufWriter, Read, Write};
use std::sync::{Arc, Mutex};

use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::vbobject::VBScriptObject;

#[derive(Debug)]
struct TextStreamInner {
    reader: Option<BufReader<File>>,
    writer: Option<BufWriter<File>>,
    line: i64,
    column: i64,
    at_end_of_stream: bool,
    at_end_of_line: bool,
    closed: bool,
}

#[derive(Debug, Clone)]
pub struct TextStream {
    inner: Arc<Mutex<TextStreamInner>>,
}

impl TextStream {
    pub fn new_read(file: File) -> Self {
        let len = file.metadata().map(|m| m.len()).unwrap_or(0);
        TextStream {
            inner: Arc::new(Mutex::new(TextStreamInner {
                reader: Some(BufReader::new(file)),
                writer: None,
                line: 1,
                column: 1,
                at_end_of_stream: len == 0,
                at_end_of_line: false,
                closed: false,
            })),
        }
    }

    pub fn new_write(file: File) -> Self {
        TextStream {
            inner: Arc::new(Mutex::new(TextStreamInner {
                reader: None,
                writer: Some(BufWriter::new(file)),
                line: 0,
                column: 0,
                at_end_of_stream: true,
                at_end_of_line: false,
                closed: false,
            })),
        }
    }

    pub fn new_append(file: File) -> Self {
        TextStream {
            inner: Arc::new(Mutex::new(TextStreamInner {
                reader: None,
                writer: Some(BufWriter::new(file)),
                line: 0,
                column: 0,
                at_end_of_stream: true,
                at_end_of_line: false,
                closed: false,
            })),
        }
    }

    fn check_closed(inner: &TextStreamInner) -> Result<(), VBSError> {
        if inner.closed {
            return Err(VBSErrorType::RuntimeError
                .into_error("TextStream is closed".to_string()));
        }
        Ok(())
    }

    fn check_reader(inner: &TextStreamInner) -> Result<(), VBSError> {
        if inner.reader.is_none() {
            return Err(VBSErrorType::RuntimeError
                .into_error("TextStream is not open for reading".to_string()));
        }
        Ok(())
    }

    fn check_writer(inner: &TextStreamInner) -> Result<(), VBSError> {
        if inner.writer.is_none() {
            return Err(VBSErrorType::RuntimeError
                .into_error("TextStream is not open for writing".to_string()));
        }
        Ok(())
    }
}

impl VBScriptObject for TextStream {
    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(&self, name: &str, _context: &mut ExecutionContext) -> Result<VBValue, VBSError> {
        let inner = self.inner.lock().unwrap();
        Self::check_closed(&inner)?;
        match name.to_uppercase().as_str() {
            "ATENDOFSTREAM" => Ok(VBValue::Boolean(inner.at_end_of_stream)),
            "ATENDOFLINE" => Ok(VBValue::Boolean(inner.at_end_of_line)),
            "LINE" => Ok(VBValue::Number(inner.line as f64)),
            "COLUMN" => Ok(VBValue::Number(inner.column as f64)),
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Property '{}' not found on TextStream", name),
            )),
        }
    }

    fn set_property(&mut self, name: &str, _value: VBValue, _context: &mut ExecutionContext) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError.into_error(
            format!("Cannot set property '{}' on TextStream", name),
        ))
    }

    fn call_method(&mut self, name: &str, args: &[VBValue]) -> Result<VBValue, VBSError> {
        let mut inner = self.inner.lock().unwrap();
        Self::check_closed(&inner)?;

        match name.to_uppercase().as_str() {
            "READ" => {
                Self::check_reader(&inner)?;
                let n = to_arg_f64(&args[0]) as usize;
                let mut buf = vec![0u8; n];
                if let Some(reader) = &mut inner.reader {
                    let bytes_read = reader.read(&mut buf).unwrap_or(0);
                    buf.truncate(bytes_read);
                    inner.at_end_of_stream = bytes_read < n;
                    for &b in &buf {
                        if b == b'\n' {
                            inner.line += 1;
                            inner.column = 1;
                        } else if b == b'\r' {
                        } else {
                            inner.column += 1;
                        }
                    }
                    inner.at_end_of_line = buf.last() == Some(&b'\n');
                    Ok(VBValue::String(String::from_utf8_lossy(&buf).to_string()))
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for reading".to_string()))
                }
            }
            "READLINE" => {
                Self::check_reader(&inner)?;
                if let Some(reader) = &mut inner.reader {
                    let mut line = String::new();
                    let bytes_read = reader.read_line(&mut line).unwrap_or(0);
                    if bytes_read == 0 {
                        inner.at_end_of_stream = true;
                        inner.at_end_of_line = true;
                        return Ok(VBValue::Empty);
                    }
                    if line.ends_with('\n') {
                        line.pop();
                        if line.ends_with('\r') {
                            line.pop();
                        }
                        inner.at_end_of_line = true;
                    }
                    inner.line += 1;
                    inner.column = 1;
                    Ok(VBValue::String(line))
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for reading".to_string()))
                }
            }
            "READALL" => {
                Self::check_reader(&inner)?;
                if let Some(reader) = &mut inner.reader {
                    let mut content = String::new();
                    reader.read_to_string(&mut content).unwrap_or(0);
                    inner.at_end_of_stream = true;
                    inner.at_end_of_line = true;
                    Ok(VBValue::String(content))
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for reading".to_string()))
                }
            }
            "WRITE" => {
                Self::check_writer(&inner)?;
                let text = to_arg_string(&args[0]);
                if let Some(writer) = &mut inner.writer {
                    let _ = write!(writer, "{}", text);
                    if let Some(last_nl) = text.rfind('\n') {
                        inner.line += text.matches('\n').count() as i64;
                        inner.column = (text.len() - last_nl) as i64;
                    } else {
                        inner.column += text.len() as i64;
                    }
                    Ok(VBValue::Empty)
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for writing".to_string()))
                }
            }
            "WRITELINE" => {
                Self::check_writer(&inner)?;
                let text = if args.is_empty() {
                    String::new()
                } else {
                    to_arg_string(&args[0])
                };
                if let Some(writer) = &mut inner.writer {
                    let _ = writeln!(writer, "{}", text);
                    inner.line += 1;
                    inner.column = 1;
                    Ok(VBValue::Empty)
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for writing".to_string()))
                }
            }
            "WRITEBLANKLINES" => {
                Self::check_writer(&inner)?;
                let n = to_arg_f64(&args[0]) as usize;
                if let Some(writer) = &mut inner.writer {
                    for _ in 0..n {
                        let _ = writeln!(writer);
                    }
                    inner.line += n as i64;
                    inner.column = 1;
                    Ok(VBValue::Empty)
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for writing".to_string()))
                }
            }
            "SKIP" => {
                Self::check_reader(&inner)?;
                let n = to_arg_f64(&args[0]) as u64;
                if let Some(reader) = &mut inner.reader {
                    let skipped =
                        std::io::copy(&mut reader.take(n), &mut std::io::sink()).unwrap_or(0);
                    inner.column += skipped as i64;
                    if skipped < n {
                        inner.at_end_of_stream = true;
                    }
                    Ok(VBValue::Empty)
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for reading".to_string()))
                }
            }
            "SKIPLINE" => {
                Self::check_reader(&inner)?;
                if let Some(reader) = &mut inner.reader {
                    let mut line = String::new();
                    let bytes = reader.read_line(&mut line).unwrap_or(0);
                    if bytes == 0 {
                        inner.at_end_of_stream = true;
                    }
                    inner.line += 1;
                    inner.column = 1;
                    Ok(VBValue::Empty)
                } else {
                    Err(VBSErrorType::RuntimeError
                        .into_error("TextStream is not open for reading".to_string()))
                }
            }
            "CLOSE" => {
                if let Some(mut writer) = inner.writer.take() {
                    writer.flush().unwrap_or(());
                }
                inner.reader = None;
                inner.closed = true;
                Ok(VBValue::Empty)
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(
                format!("Method '{}' not found on TextStream", name),
            )),
        }
    }

    fn indexed_get(&self, _index: &VBValue) -> Result<VBValue, VBSError> {
        Err(VBSErrorType::RuntimeError
            .into_error("TextStream does not support indexed access".to_string()))
    }

    fn indexed_set(&mut self, _index: &VBValue, _value: VBValue) -> Result<(), VBSError> {
        Err(VBSErrorType::RuntimeError
            .into_error("TextStream does not support indexed access".to_string()))
    }
}

fn to_arg_string(val: &VBValue) -> String {
    match val {
        VBValue::String(s) => s.clone(),
        VBValue::Null => "Null".to_string(),
        VBValue::Empty => "".to_string(),
        VBValue::Number(n) => n.to_string(),
        VBValue::Boolean(true) => "True".to_string(),
        VBValue::Boolean(false) => "False".to_string(),
        VBValue::Array(_) => "Array".to_string(),
        VBValue::Object(_) => "Object".to_string(),
    }
}

fn to_arg_f64(val: &VBValue) -> f64 {
    match val {
        VBValue::Number(n) => *n,
        VBValue::String(s) => s.parse::<f64>().unwrap_or(0.0),
        VBValue::Boolean(true) => -1.0,
        VBValue::Boolean(false) => 0.0,
        VBValue::Null | VBValue::Empty | VBValue::Array(_) | VBValue::Object(_) => 0.0,
    }
}
