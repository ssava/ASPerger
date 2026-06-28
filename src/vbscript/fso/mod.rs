pub(crate) use self::file::FileObject;
pub(crate) use self::filesystem::FileSystemObject;
pub(crate) use self::folder::FolderObject;

mod file;
mod filesystem;
mod folder;

fn resolve_path(path: &str) -> std::path::PathBuf {
    let p = std::path::Path::new(path);
    if p.is_absolute() {
        p.to_path_buf()
    } else {
        let cwd = std::env::current_dir().unwrap_or_default();
        cwd.join(p)
    }
}

fn format_datetime(t: Option<std::time::SystemTime>) -> String {
    match t {
        Some(time) => {
            let duration = time
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap_or_default();
            let secs = duration.as_secs();
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

fn infer_type(path: &std::path::Path) -> String {
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

fn copy_dir_recursive(src: &std::path::Path, dst: &std::path::Path) -> std::io::Result<()> {
    if !dst.exists() {
        std::fs::create_dir_all(dst)?;
    }
    for entry in std::fs::read_dir(src)? {
        let entry = entry?;
        let entry_path = entry.path();
        let dest_path = dst.join(entry.file_name());
        if entry_path.is_dir() {
            copy_dir_recursive(&entry_path, &dest_path)?;
        } else {
            std::fs::copy(&entry_path, &dest_path)?;
        }
    }
    Ok(())
}
