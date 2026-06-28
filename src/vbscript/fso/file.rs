use std::fs;
use std::path::{Path, PathBuf};

use super::super::execution_context::ExecutionContext;
use super::super::textstream::TextStream;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::{format_datetime, infer_type};
use crate::vbscript::vbobject::VBScriptObject;
use crate::{prop_not_found, method_not_found, cannot_set_property};

#[derive(Debug)]
pub struct FileObject {
    path: PathBuf,
    name: String,
    short_name: String,
    size: u64,
    date_created: String,
    date_last_modified: String,
    date_last_accessed: String,
    file_type: String,
    parent_folder: String,
}

impl FileObject {
    pub fn new(path: &Path) -> Self {
        let meta = fs::metadata(path).ok();
        let size = meta.as_ref().map(|m| m.len()).unwrap_or(0);
        let created = meta.as_ref().and_then(|m| m.created().ok());
        let modified = meta.as_ref().and_then(|m| m.modified().ok());
        let accessed = meta.as_ref().and_then(|m| m.accessed().ok());

        FileObject {
            path: path.to_path_buf(),
            name: path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("")
                .to_string(),
            short_name: path
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("")
                .to_string(),
            size,
            date_created: format_datetime(created),
            date_last_modified: format_datetime(modified),
            date_last_accessed: format_datetime(accessed),
            file_type: infer_type(path),
            parent_folder: path
                .parent()
                .and_then(|p| p.to_str())
                .unwrap_or("")
                .to_string(),
        }
    }
}

impl VBScriptObject for FileObject {
    fn type_name(&self) -> &'static str {
        "File"
    }

    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(FileObject {
            path: self.path.clone(),
            name: self.name.clone(),
            short_name: self.short_name.clone(),
            size: self.size,
            date_created: self.date_created.clone(),
            date_last_modified: self.date_last_modified.clone(),
            date_last_accessed: self.date_last_accessed.clone(),
            file_type: self.file_type.clone(),
            parent_folder: self.parent_folder.clone(),
        })
    }

    fn get_property(
        &self,
        name: &str,
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "NAME" => Ok(VBValue::String(self.name.clone())),
            "SHORTPATH" | "PATH" => Ok(VBValue::String(
                self.path.to_str().unwrap_or("").to_string(),
            )),
            "SHORTNAME" => Ok(VBValue::String(self.short_name.clone())),
            "SIZE" => Ok(VBValue::Number(self.size as f64)),
            "TYPE" => Ok(VBValue::String(self.file_type.clone())),
            "DATECREATED" => Ok(VBValue::String(self.date_created.clone())),
            "DATELASTMODIFIED" => Ok(VBValue::String(self.date_last_modified.clone())),
            "DATELASTACCESSED" => Ok(VBValue::String(self.date_last_accessed.clone())),
            "PARENTFOLDER" => Ok(VBValue::String(self.parent_folder.clone())),
            "ATTRIBUTES" => {
                let attrs = fs::metadata(&self.path)
                    .map(|m| {
                        let mut a = 0i32;
                        if m.permissions().readonly() {
                            a += 1;
                        }
                        if m.is_dir() {
                            a += 16;
                        }
                        a
                    })
                    .unwrap_or(0);
                Ok(VBValue::Number(attrs as f64))
            }
            _ => prop_not_found!("File", name),
        }
    }

    fn set_property(
        &mut self,
        name: &str,
        value: VBValue,
        _context: &mut ExecutionContext,
    ) -> Result<(), VBSError> {
        match name.to_uppercase().as_str() {
            "NAME" => {
                let new_name = value_utils::to_arg_string(&value);
                let new_path = self.path.with_file_name(&new_name);
                fs::rename(&self.path, &new_path).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!("Cannot rename file: {}", e))
                })?;
                self.path = new_path;
                self.name = new_name;
                Ok(())
            }
            _ => cannot_set_property!("File", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "DELETE" => {
                let force = if !args.is_empty() {
                    value_utils::to_boolean(&args[0])
                } else {
                    false
                };
                if force {
                    fs::remove_file(&self.path).map_err(|e| {
                        VBSErrorType::RuntimeError.into_error(format!("Cannot delete file: {}", e))
                    })?;
                } else {
                    fs::remove_file(&self.path).map_err(|e| {
                        VBSErrorType::RuntimeError.into_error(format!("Cannot delete file: {}", e))
                    })?;
                }
                Ok(VBValue::Empty)
            }
            "OPENASTEXTSTREAM" => {
                let iomode = if !args.is_empty() {
                    value_utils::to_arg_f64(&args[0]) as i32
                } else {
                    1
                };
                let file = match iomode {
                    1 => fs::File::open(&self.path),
                    2 => fs::File::create(&self.path),
                    8 => fs::OpenOptions::new().append(true).open(&self.path),
                    _ => {
                        return Err(VBSErrorType::RuntimeError
                            .into_error(format!("Invalid IOMode: {}", iomode)))
                    }
                }
                .map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!("Cannot open file: {}", e))
                })?;
                let ts = match iomode {
                    1 => TextStream::new_read(file),
                    2 => TextStream::new_write(file),
                    _ => TextStream::new_append(file),
                };
                Ok(VBValue::Object(Box::new(ts)))
            }
            _ => method_not_found!("File", name),
        }
    }
}
