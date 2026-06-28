use std::fs;

use super::super::execution_context::ExecutionContext;
use super::super::textstream::TextStream;
use super::super::value::VBValue;
use super::super::value_utils;
use super::super::vbs_error::{VBSError, VBSErrorType};
use super::{copy_dir_recursive, resolve_path};
use super::{FileObject, FolderObject};
use crate::vbscript::vbobject::VBScriptObject;
use crate::{prop_not_found, method_not_found};

#[derive(Debug, Clone, Default)]
pub struct FileSystemObject;

impl FileSystemObject {
    pub fn new() -> Self {
        FileSystemObject
    }
}

impl VBScriptObject for FileSystemObject {
    fn type_name(&self) -> &'static str {
        "FileSystemObject"
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
            "DRIVES" => {
                #[cfg(unix)]
                {
                    let drives = vec![VBValue::String("/".to_string())];
                    Ok(VBValue::Array(std::sync::Arc::new(drives), vec![]))
                }
                #[cfg(windows)]
                {
                    let drives: Vec<VBValue> = (b'A'..=b'Z')
                        .filter_map(|letter| {
                            let path = format!("{}:\\", letter as char);
                            if Path::new(&path).exists() {
                                Some(VBValue::String(path))
                            } else {
                                None
                            }
                        })
                        .collect();
                    Ok(VBValue::Array(std::sync::Arc::new(drives), vec![]))
                }
            }
            _ => prop_not_found!("FileSystemObject", name),
        }
    }

    fn call_method(
        &mut self,
        name: &str,
        args: &[VBValue],
        _context: &mut ExecutionContext,
    ) -> Result<VBValue, VBSError> {
        match name.to_uppercase().as_str() {
            "FILEEXISTS" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                Ok(VBValue::Boolean(p.is_file()))
            }
            "FOLDEREXISTS" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                Ok(VBValue::Boolean(p.is_dir()))
            }
            "GETFILE" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                if !p.is_file() {
                    return Err(
                        VBSErrorType::RuntimeError.into_error(format!("File not found: {}", path))
                    );
                }
                let file = FileObject::new(&p);
                Ok(VBValue::Object(Box::new(file)))
            }
            "GETFOLDER" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                if !p.is_dir() {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("Folder not found: {}", path)));
                }
                let folder = FolderObject::new(&p);
                Ok(VBValue::Object(Box::new(folder)))
            }
            "CREATETEXTFILE" => {
                let path = value_utils::to_arg_string(&args[0]);
                let overwrite = if args.len() > 1 {
                    value_utils::to_boolean(&args[1])
                } else {
                    false
                };
                let p = resolve_path(&path);
                if p.exists() && !overwrite {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("File already exists: {}", path)));
                }
                if let Some(parent) = p.parent() {
                    fs::create_dir_all(parent).unwrap_or(());
                }
                let file = fs::File::create(&p).map_err(|e| {
                    VBSErrorType::RuntimeError
                        .into_error(format!("Cannot create file '{}': {}", path, e))
                })?;
                Ok(VBValue::Object(Box::new(TextStream::new_write(file))))
            }
            "OPENTEXTFILE" => {
                let path = value_utils::to_arg_string(&args[0]);
                let iomode = if args.len() > 1 {
                    value_utils::to_arg_f64(&args[1]) as i32
                } else {
                    1
                };
                let create = if args.len() > 2 {
                    value_utils::to_boolean(&args[2])
                } else {
                    false
                };
                let p = resolve_path(&path);
                if !p.exists() {
                    if create {
                        if let Some(parent) = p.parent() {
                            fs::create_dir_all(parent).unwrap_or(());
                        }
                        let file = fs::File::create(&p).map_err(|e| {
                            VBSErrorType::RuntimeError
                                .into_error(format!("Cannot create file '{}': {}", path, e))
                        })?;
                        if iomode == 1 {
                            drop(file);
                            let file2 = fs::File::open(&p).map_err(|_| {
                                VBSErrorType::RuntimeError
                                    .into_error(format!("Cannot open file: {}", path))
                            })?;
                            Ok(VBValue::Object(Box::new(TextStream::new_read(file2))))
                        } else if iomode == 2 {
                            Ok(VBValue::Object(Box::new(TextStream::new_write(file))))
                        } else {
                            Ok(VBValue::Object(Box::new(TextStream::new_append(file))))
                        }
                    } else {
                        Err(VBSErrorType::RuntimeError
                            .into_error(format!("File not found: {}", path)))
                    }
                } else {
                    match iomode {
                        1 => {
                            let file = fs::File::open(&p).map_err(|_| {
                                VBSErrorType::RuntimeError
                                    .into_error(format!("Cannot open file: {}", path))
                            })?;
                            Ok(VBValue::Object(Box::new(TextStream::new_read(file))))
                        }
                        2 => {
                            let file = fs::File::create(&p).map_err(|_| {
                                VBSErrorType::RuntimeError
                                    .into_error(format!("Cannot open file: {}", path))
                            })?;
                            Ok(VBValue::Object(Box::new(TextStream::new_write(file))))
                        }
                        8 => {
                            let file =
                                fs::OpenOptions::new().append(true).open(&p).map_err(|_| {
                                    VBSErrorType::RuntimeError
                                        .into_error(format!("Cannot open file: {}", path))
                                })?;
                            Ok(VBValue::Object(Box::new(TextStream::new_append(file))))
                        }
                        _ => Err(VBSErrorType::RuntimeError
                            .into_error(format!("Invalid IOMode: {}", iomode))),
                    }
                }
            }
            "DELETEFILE" => {
                let path = value_utils::to_arg_string(&args[0]);
                let force = if args.len() > 1 {
                    value_utils::to_boolean(&args[1])
                } else {
                    false
                };
                let p = resolve_path(&path);
                if p.is_file() {
                    fs::remove_file(&p).map_err(|e| {
                        VBSErrorType::RuntimeError
                            .into_error(format!("Cannot delete file '{}': {}", path, e))
                    })?;
                } else if !force {
                    return Err(
                        VBSErrorType::RuntimeError.into_error(format!("File not found: {}", path))
                    );
                }
                Ok(VBValue::Empty)
            }
            "DELETEFOLDER" => {
                let path = value_utils::to_arg_string(&args[0]);
                let force = if args.len() > 1 {
                    value_utils::to_boolean(&args[1])
                } else {
                    false
                };
                let p = resolve_path(&path);
                if p.is_dir() {
                    if force {
                        fs::remove_dir_all(&p).map_err(|e| {
                            VBSErrorType::RuntimeError
                                .into_error(format!("Cannot delete folder '{}': {}", path, e))
                        })?;
                    } else {
                        fs::remove_dir(&p).map_err(|e| {
                            VBSErrorType::RuntimeError
                                .into_error(format!("Cannot delete folder '{}': {}", path, e))
                        })?;
                    }
                } else if !force {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("Folder not found: {}", path)));
                }
                Ok(VBValue::Empty)
            }
            "COPYFILE" => {
                let source = value_utils::to_arg_string(&args[0]);
                let dest = value_utils::to_arg_string(&args[1]);
                let overwrite = if args.len() > 2 {
                    value_utils::to_boolean(&args[2])
                } else {
                    true
                };
                let src = resolve_path(&source);
                let dst = resolve_path(&dest);
                if dst.exists() && !overwrite {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("Destination file already exists: {}", dest)));
                }
                if let Some(parent) = dst.parent() {
                    fs::create_dir_all(parent).unwrap_or(());
                }
                fs::copy(&src, &dst).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!(
                        "Cannot copy file '{}' to '{}': {}",
                        source, dest, e
                    ))
                })?;
                Ok(VBValue::Empty)
            }
            "COPYFOLDER" => {
                let source = value_utils::to_arg_string(&args[0]);
                let dest = value_utils::to_arg_string(&args[1]);
                let overwrite = if args.len() > 2 {
                    value_utils::to_boolean(&args[2])
                } else {
                    true
                };
                let src = resolve_path(&source);
                let dst = resolve_path(&dest);
                if dst.exists() && !overwrite {
                    return Err(VBSErrorType::RuntimeError
                        .into_error(format!("Destination folder already exists: {}", dest)));
                }
                copy_dir_recursive(&src, &dst).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!(
                        "Cannot copy folder '{}' to '{}': {}",
                        source, dest, e
                    ))
                })?;
                Ok(VBValue::Empty)
            }
            "MOVEFILE" => {
                let source = value_utils::to_arg_string(&args[0]);
                let dest = value_utils::to_arg_string(&args[1]);
                let src = resolve_path(&source);
                let dst = resolve_path(&dest);
                if let Some(parent) = dst.parent() {
                    fs::create_dir_all(parent).unwrap_or(());
                }
                fs::rename(&src, &dst).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!(
                        "Cannot move file '{}' to '{}': {}",
                        source, dest, e
                    ))
                })?;
                Ok(VBValue::Empty)
            }
            "MOVEFOLDER" => {
                let source = value_utils::to_arg_string(&args[0]);
                let dest = value_utils::to_arg_string(&args[1]);
                let src = resolve_path(&source);
                let dst = resolve_path(&dest);
                if let Some(parent) = dst.parent() {
                    fs::create_dir_all(parent).unwrap_or(());
                }
                fs::rename(&src, &dst).map_err(|e| {
                    VBSErrorType::RuntimeError.into_error(format!(
                        "Cannot move folder '{}' to '{}': {}",
                        source, dest, e
                    ))
                })?;
                Ok(VBValue::Empty)
            }
            "CREATEFOLDER" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                fs::create_dir_all(&p).map_err(|e| {
                    VBSErrorType::RuntimeError
                        .into_error(format!("Cannot create folder '{}': {}", path, e))
                })?;
                Ok(VBValue::Empty)
            }
            "GETPARENTFOLDERNAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = std::path::Path::new(&path);
                let parent = p.parent().and_then(|p| {
                    let s = p.to_str()?;
                    if s.is_empty() {
                        None
                    } else {
                        Some(s.to_string())
                    }
                });
                match parent {
                    Some(name) => Ok(VBValue::String(name)),
                    None => Ok(VBValue::Empty),
                }
            }
            "GETFILENAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = std::path::Path::new(&path);
                let name = p.file_name().and_then(|n| n.to_str());
                match name {
                    Some(n) => Ok(VBValue::String(n.to_string())),
                    None => Ok(VBValue::Empty),
                }
            }
            "GETEXTENSIONNAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = std::path::Path::new(&path);
                let ext = p.extension().and_then(|e| e.to_str());
                match ext {
                    Some(e) => Ok(VBValue::String(e.to_string())),
                    None => Ok(VBValue::String("".to_string())),
                }
            }
            "GETBASENAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = std::path::Path::new(&path);
                let name = p.file_stem().and_then(|n| n.to_str());
                match name {
                    Some(n) => Ok(VBValue::String(n.to_string())),
                    None => Ok(VBValue::String("".to_string())),
                }
            }
            "GETABSOLUTEPATHNAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = resolve_path(&path);
                let s = p.to_str().unwrap_or(&path);
                Ok(VBValue::String(s.to_string()))
            }
            "BUILDPATH" => {
                let parent = value_utils::to_arg_string(&args[0]);
                let child = value_utils::to_arg_string(&args[1]);
                let mut parent = parent.trim_end_matches('/').trim_end_matches('\\');
                if parent.is_empty() {
                    parent = ".";
                }
                let child = child.trim_start_matches('/').trim_start_matches('\\');
                let joined = format!("{}/{}", parent, child);
                let p = std::path::Path::new(&joined);
                let result = p.to_str().unwrap_or(&joined);
                Ok(VBValue::String(result.to_string()))
            }
            "GETSPECIALFOLDER" => {
                let folder_type = value_utils::to_arg_f64(&args[0]) as i32;
                match folder_type {
                    0 => Ok(VBValue::String(
                        std::env::var("WINDIR").unwrap_or_else(|_| "/usr".to_string()),
                    )),
                    1 => Ok(VBValue::String(
                        std::env::var("SYSTEMROOT")
                            .or_else(|_| std::env::var("WINDIR"))
                            .unwrap_or_else(|_| "/usr/lib".to_string()),
                    )),
                    2 => Ok(VBValue::String(
                        std::env::temp_dir().to_str().unwrap_or("/tmp").to_string(),
                    )),
                    _ => Err(VBSErrorType::RuntimeError
                        .into_error(format!("Invalid SpecialFolder constant: {}", folder_type))),
                }
            }
            _ => method_not_found!("FileSystemObject", name),
        }
    }
}
