//! ADODB COM object stubs: `Connection`, `Recordset`, `Field`, `Parameter`,
//! and related types. Provides minimal script-level access for basic database
//! scenarios.

use super::execution_context::ExecutionContext;
use super::value::VBValue;
use super::value_utils;
use super::vbobject::VBScriptObject;
use super::vbs_error::{VBSError, VBSErrorType};

// ---- Connection ----

#[derive(Debug, Clone)]
/// `ADODB.Connection` — database connection stub.
///
/// Provides basic `Open`, `Close`, `Execute` methods and a `ConnectionString`
/// / `State` property.  Currently stubbed with minimal real DB integration.
pub struct Connection {
    connection_string: String,
    state: i32,
}

impl Default for Connection {
    fn default() -> Self {
        Self::new()
    }
}

impl Connection {
    pub fn new() -> Self {
        Connection {
            connection_string: String::new(),
            state: 0,
        }
    }
}

impl VBScriptObject for Connection {
    fn type_name(&self) -> &'static str {
        "Connection"
    }

    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "CONNECTIONSTRING" => Ok(VBValue::String(self.connection_string.clone())),
            "STATE" => Ok(VBValue::Number(self.state as f64)),
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Property '{}' not found on Connection object",
                name
            ))),
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            "CONNECTIONSTRING" => {
                self.connection_string = value_utils::to_arg_string(&value);
                Ok(())
            }
            _ => Err(VBSErrorType::RuntimeError.into_error(format!(
                "Cannot set property '{}' on Connection object",
                name
            ))),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "OPEN" => {
                if !args.is_empty() {
                    self.connection_string = value_utils::to_arg_string(&args[0]);
                }
                self.state = 1;
                Ok(VBValue::Empty)
            }
            "CLOSE" => {
                self.state = 0;
                Ok(VBValue::Empty)
            }
            "EXECUTE" => {
                if args.is_empty() {
                    return Err(VBSErrorType::ValueError.into_error(
                        "Connection.Execute requires at least 1 argument (sql)".to_string(),
                    ));
                }
                let _sql = value_utils::to_arg_string(&args[0]);
                Ok(VBValue::Object(Box::new(Recordset::empty())))
            }
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Method '{}' not found on Connection object", name))),
        }
    }
}

// ---- Recordset ----

#[derive(Debug, Clone)]
pub struct Recordset {
    eof: bool,
    field_names: Vec<String>,
    current_index: usize,
}

impl Recordset {
    pub fn empty() -> Self {
        Recordset {
            eof: true,
            field_names: Vec::new(),
            current_index: 0,
        }
    }
}

impl VBScriptObject for Recordset {
    fn type_name(&self) -> &'static str {
        "Recordset"
    }

    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(self.clone())
    }

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "EOF" => Ok(VBValue::Boolean(self.eof)),
            "RECORDCOUNT" => Ok(VBValue::Number(self.field_names.len() as f64)),
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Property '{}' not found on Recordset object", name))),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        _args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "MOVENEXT" => {
                if self.eof {
                    return Err(
                        VBSErrorType::RuntimeError.into_error("Cannot move past EOF".to_string())
                    );
                }
                self.current_index += 1;
                if self.current_index >= self.field_names.len() {
                    self.eof = true;
                }
                Ok(VBValue::Empty)
            }
            "CLOSE" => {
                self.eof = true;
                self.current_index = 0;
                Ok(VBValue::Empty)
            }
            _ => Err(VBSErrorType::RuntimeError
                .into_error(format!("Method '{}' not found on Recordset object", name))),
        }
    }
}
