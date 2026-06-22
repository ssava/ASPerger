use std::path::Path;

use clap::Parser;

/// ASP server CLI configuration.
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
pub struct Config {
    /// Host address the server will listen on.
    #[clap(long, env = "ASPERGER_HOST", default_value = "127.0.0.1")]
    pub host: String,

    /// Port the server will listen on.
    #[clap(short, long, env = "ASPERGER_PORT", default_value = "8080")]
    pub port: u16,

    /// Directory containing ASP files to serve.
    #[clap(short, long, env = "ASPERGER_FOLDER", default_value = ".")]
    pub folder: String,

    /// Path to an .asp file or directory (positional shortcut for --folder).
    pub program: Option<String>,

    /// Enable directory listing when no default document is found.
    #[clap(long, env = "ASPERGER_DIRECTORY_LISTING")]
    pub enable_directory_listing: bool,
}

/// Runtime server configuration with all overrides applied in order:
/// defaults < INI file < programmatic overrides.
#[derive(Debug, Clone)]
pub struct AspServerConfig {
    pub host: String,
    pub port: u16,
    pub folder: String,
    pub default_document: String,
    pub directory_listing: bool,
}

impl Default for AspServerConfig {
    fn default() -> Self {
        Self {
            host: "127.0.0.1".to_string(),
            port: 9090,
            folder: ".".to_string(),
            default_document: "index.asp".to_string(),
            directory_listing: false,
        }
    }
}

impl AspServerConfig {
    /// Load `asp.ini` from the served folder and apply its values on top of defaults.
    pub fn from_folder(folder: &str) -> Self {
        let mut cfg = Self { folder: folder.to_string(), ..Self::default() };

        let ini_path = Path::new(folder).join("asp.ini");
        if let Ok(content) = std::fs::read_to_string(&ini_path) {
            let mut in_server = false;
            for line in content.lines() {
                let line = line.trim();
                if line.starts_with('[') && line.ends_with(']') {
                    in_server = line.eq_ignore_ascii_case("[server]");
                    continue;
                }
                if !in_server || line.is_empty() || line.starts_with('#') || line.starts_with(';') {
                    continue;
                }
                if let Some((key, value)) = line.split_once('=') {
                    let key = key.trim().to_lowercase();
                    let value = value.trim();
                    match key.as_str() {
                        "host" => cfg.host = value.to_string(),
                        "port" => {
                            if let Ok(p) = value.parse::<u16>() {
                                cfg.port = p;
                            }
                        }
                        "default_document" => cfg.default_document = value.to_string(),
                        "enable_directory_listing" => {
                            cfg.directory_listing = value.eq_ignore_ascii_case("true");
                        }
                        _ => {}
                    }
                }
            }
        }

        cfg
    }

    /// Apply overrides from a map (e.g. DAP launch args or CLI args).
    /// Empty/Never values are ignored so INI/defaults are preserved.
    pub fn apply_overrides(&mut self, host: Option<&str>, port: Option<u16>, folder: Option<&str>, default_document: Option<&str>, directory_listing: Option<bool>) {
        if let Some(h) = host {
            if !h.is_empty() {
                self.host = h.to_string();
            }
        }
        if let Some(p) = port {
            self.port = p;
        }
        if let Some(f) = folder {
            if !f.is_empty() {
                self.folder = f.to_string();
            }
        }
        if let Some(d) = default_document {
            if !d.is_empty() {
                self.default_document = d.to_string();
            }
        }
        if let Some(dl) = directory_listing {
            self.directory_listing = dl;
        }
    }
}
