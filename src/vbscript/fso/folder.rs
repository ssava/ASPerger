use std::fs;
use std::path::{Path, PathBuf};

use super::super::execution_context::ExecutionContext;
use super::super::textstream::TextStream;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::{format_datetime, FileObject};
use crate::vbscript::vbobject::VBScriptObject;
use crate::{prop_not_found, method_not_found, cannot_set_property};

#[derive(Debug)]
pub struct FolderObject {
    path: PathBuf,
    name: String,
    size: u64,
    date_created: String,
    date_last_modified: String,
    date_last_accessed: String,
    is_root: bool,
    parent_folder: String,
}

impl FolderObject {
    pub fn new(path: &Path) -> Self {
        let meta = fs::metadata(path).ok();
        let size = meta.as_ref().map(|m| m.len()).unwrap_or(0);
        let created = meta.as_ref().and_then(|m| m.created().ok());
        let modified = meta.as_ref().and_then(|m| m.modified().ok());
        let accessed = meta.as_ref().and_then(|m| m.accessed().ok());

        let name = path
            .file_name()
            .and_then(|n| n.to_str())
            .map(|s| s.to_string())
            .unwrap_or_else(|| path.to_str().unwrap_or("").to_string());

        FolderObject {
            path: path.to_path_buf(),
            name,
            size,
            date_created: format_datetime(created),
            date_last_modified: format_datetime(modified),
            date_last_accessed: format_datetime(accessed),
            is_root: path.parent().is_none() || path == Path::new("/"),
            parent_folder: path
                .parent()
                .and_then(|p| p.to_str())
                .unwrap_or("")
                .to_string(),
        }
    }
}

impl VBScriptObject for FolderObject {
    fn type_name(&self) -> &'static str {
        "Folder"
    }

    fn clone_box(&self) -> Box<dyn VBScriptObject> {
        Box::new(FolderObject {
            path: self.path.clone(),
            name: self.name.clone(),
            size: self.size,
            date_created: self.date_created.clone(),
            date_last_modified: self.date_last_modified.clone(),
            date_last_accessed: self.date_last_accessed.clone(),
            is_root: self.is_root,
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
            "SHORTNAME" => Ok(VBValue::String(self.name.clone())),
            "SIZE" => Ok(VBValue::Number(self.size as f64)),
            "TYPE" => Ok(VBValue::String("File Folder".to_string())),
            "DATECREATED" => Ok(VBValue::String(self.date_created.clone())),
            "DATELASTMODIFIED" => Ok(VBValue::String(self.date_last_modified.clone())),
            "DATELASTACCESSED" => Ok(VBValue::String(self.date_last_accessed.clone())),
            "ISROOTFOLDER" => Ok(VBValue::Boolean(self.is_root)),
            "PARENTFOLDER" => Ok(VBValue::String(self.parent_folder.clone())),
            "FILES" => {
                let entries: Vec<VBValue> = match fs::read_dir(&self.path) {
                    Ok(dir) => dir
                        .filter_map(|e| e.ok())
                        .filter(|e| e.path().is_file())
                        .map(|e| VBValue::Object(Box::new(FileObject::new(&e.path()))))
                        .collect(),
                    Err(_) => vec![],
                };
                Ok(VBValue::Array(std::sync::Arc::new(entries), vec![]))
            }
            "SUBFOLDERS" => {
                let entries: Vec<VBValue> = match fs::read_dir(&self.path) {
                    Ok(dir) => dir
                        .filter_map(|e| e.ok())
                        .filter(|e| e.path().is_dir())
                        .map(|e| VBValue::Object(Box::new(FolderObject::new(&e.path()))))
                        .collect(),
                    Err(_) => vec![],
                };
                Ok(VBValue::Array(std::sync::Arc::new(entries), vec![]))
            }
            "ATTRIBUTES" => {
                let a = 16i32;
                Ok(VBValue::Number(a as f64))
            }
            _ => prop_not_found!("Folder", name),
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
                    VBSErrorType::RuntimeError.into_error(format!("Cannot rename folder: {}", e))
                })?;
                self.path = new_path;
                self.name = new_name;
                Ok(())
            }
            _ => cannot_set_property!("Folder", name),
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
                    fs::remove_dir_all(&self.path).map_err(|e| {
                        VBSErrorType::RuntimeError
                            .into_error(format!("Cannot delete folder: {}", e))
                    })?;
                } else {
                    fs::remove_dir(&self.path).map_err(|e| {
                        VBSErrorType::RuntimeError
                            .into_error(format!("Cannot delete folder: {}", e))
                    })?;
                }
                Ok(VBValue::Empty)
            }
            "CREATETEXTFILE" => {
                let path = value_utils::to_arg_string(&args[0]);
                let overwrite = if args.len() > 1 {
                    value_utils::to_boolean(&args[1])
                } else {
                    false
                };
                let file_path = self.path.join(&path);
                if file_path.exists() && !overwrite {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("File already exists: {}", file_path.display())));
                }
                let file = fs::File::create(&file_path).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!("Cannot create file: {}", e))
                })?;
                Ok(VBValue::Object(Box::new(TextStream::new_write(file))))
            }
            _ => method_not_found!("Folder", name),
        }
    }
}
