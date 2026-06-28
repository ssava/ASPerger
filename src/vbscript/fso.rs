//! Scripting.FileSystemObject implementation: `FileSystemObject`, `File`,
//! `Folder`, `Drive` objects with methods for file I/O and directory traversal.

use std::fs;
use std::path::{Path, PathBuf};

use super::execution_context::ExecutionContext;
use super::textstream::TextStream;
use super::value::VBValue;
use super::value_utils;
use super::vbs_error::{VBSError, VBSErrorType};
use crate::vbscript::vbobject::VBScriptObject;
use crate::{prop_not_found, method_not_found, cannot_set_property};

// ---- FileSystemObject ----

/// VBScript `Scripting.FileSystemObject` — provides file I/O,
/// directory traversal, and drive enumeration.
///
/// Methods: `CreateTextFile`, `OpenTextFile`, `GetFile`, `GetFolder`,
/// `GetDrive`, `FileExists`, `FolderExists`, `DriveExists`, `DeleteFile`,
/// `DeleteFolder`, `CopyFile`, `CopyFolder`, `MoveFile`, `MoveFolder`,
/// `CreateFolder`, `GetTempName`, `GetParentFolderName`,
/// `GetFileName`, `GetExtensionName`, `GetBaseName`, `BuildPath`.
#[derive(Debug, Clone, Default)]
pub struct FileSystemObject;

impl FileSystemObject {
    pub fn new() -> Self {
        FileSystemObject
    }
}

/// Resolve a VBScript-style path to an absolute `PathBuf`.
/// If the path is already absolute, returns it unchanged;
/// otherwise joins it with the current working directory.
fn resolve_path(path: &str) -> PathBuf {
    let p = Path::new(path);
    if p.is_absolute() {
        p.to_path_buf()
    } else {
        let cwd = std::env::current_dir().unwrap_or_default();
        cwd.join(p)
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
                            // ForReading - reopen for read
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
                let p = Path::new(&path);
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
                let p = Path::new(&path);
                let name = p.file_name().and_then(|n| n.to_str());
                match name {
                    Some(n) => Ok(VBValue::String(n.to_string())),
                    None => Ok(VBValue::Empty),
                }
            }
            "GETEXTENSIONNAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = Path::new(&path);
                let ext = p.extension().and_then(|e| e.to_str());
                match ext {
                    Some(e) => Ok(VBValue::String(e.to_string())),
                    None => Ok(VBValue::String("".to_string())),
                }
            }
            "GETBASENAME" => {
                let path = value_utils::to_arg_string(&args[0]);
                let p = Path::new(&path);
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
                let p = Path::new(&joined);
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

// ---- FileObject ----

#[derive(Debug)]
/// `Scripting.File` — represents a single file, returned by
/// `FileSystemObject.GetFile`.  Provides read-only properties
/// (`Name`, `Path`, `Size`, `DateCreated`, `DateLastModified`,
/// `DateLastAccessed`, `Type`, `ParentFolder`) and methods
/// (`Delete`, `Copy`, `Move`, `OpenAsTextStream`).
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

fn format_datetime(t: Option<std::time::SystemTime>) -> String {
    match t {
        Some(time) => {
            let duration = time
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default();
            let secs = duration.as_secs();
            // Simple UTC formatting without chrono dependency
            let days = secs / 86400;
            let time_secs = secs % 86400;
            let hours = time_secs / 3600;
            let minutes = (time_secs % 3600) / 60;
            let seconds = time_secs % 60;

            let mut year = 1970i64;
            let mut remaining_days = days as i64;

            loop {
                let days_in_year = if is_leap(year) { 366 } else { 365 };
                if remaining_days < days_in_year {
                    break;
                }
                remaining_days -= days_in_year;
                year += 1;
            }

            let month_days = if is_leap(year) {
                [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            } else {
                [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
            };

            let mut month = 1usize;
            let mut day = remaining_days;
            for (i, &md) in month_days.iter().enumerate() {
                if day < md {
                    month = i + 1;
                    break;
                }
                day -= md;
                month = i + 2;
            }
            let day = day + 1;

            format!(
                "{:02}/{:02}/{:04} {:02}:{:02}:{:02}",
                month, day, year, hours, minutes, seconds
            )
        }
        None => "".to_string(),
    }
}

fn is_leap(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn infer_type(path: &Path) -> String {
    match path.extension().and_then(|e| e.to_str()) {
        Some("txt") => "Text Document".to_string(),
        Some("html") | Some("htm") => "HTML Document".to_string(),
        Some("js") => "JavaScript File".to_string(),
        Some("css") => "CSS File".to_string(),
        Some("asp") => "ASP File".to_string(),
        Some("jpg") | Some("jpeg") => "JPEG Image".to_string(),
        Some("png") => "PNG Image".to_string(),
        Some("gif") => "GIF Image".to_string(),
        Some("pdf") => "PDF Document".to_string(),
        Some("zip") => "ZIP Archive".to_string(),
        Some("exe") => "Application".to_string(),
        Some("dll") => "Application Extension".to_string(),
        Some("xml") => "XML Document".to_string(),
        Some("json") => "JSON File".to_string(),
        _ => {
            if path.is_dir() {
                "File Folder".to_string()
            } else {
                format!(
                    "{} File",
                    path.extension()
                        .and_then(|e| e.to_str())
                        .unwrap_or("Unknown")
                )
            }
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

// ---- FolderObject ----

#[derive(Debug)]
/// `Scripting.Folder` — represents a directory, returned by
/// `FileSystemObject.GetFolder`.  Provides properties (`Name`, `Path`,
/// `Size`, `DateCreated`, `Files`, `SubFolders`, `ParentFolder`, `IsRootFolder`)
/// and methods (`Delete`, `Copy`, `Move`, `CreateTextFile`).
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

fn copy_dir_recursive(src: &Path, dst: &Path) -> std::io::Result<()> {
    if !dst.exists() {
        fs::create_dir_all(dst)?;
    }
    for entry in fs::read_dir(src)? {
        let entry = entry?;
        let entry_path = entry.path();
        let dest_path = dst.join(entry.file_name());
        if entry_path.is_dir() {
            copy_dir_recursive(&entry_path, &dest_path)?;
        } else {
            fs::copy(&entry_path, &dest_path)?;
        }
    }
    Ok(())
}
