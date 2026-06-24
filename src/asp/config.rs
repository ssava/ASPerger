use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::RwLock;

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

    /// Comma-separated list of default documents to try when requesting a directory.
    #[clap(long, env = "ASPERGER_DEFAULT_DOCUMENTS")]
    pub default_documents: Option<String>,
}

/// Per-directory settings for an ASP request.
#[derive(Debug, Clone)]
pub struct AspDirConfig {
    /// Prioritized list of default document filenames to try when a directory is requested.
    pub default_documents: Vec<String>,
    /// Whether to show directory listing when no default document is found.
    pub directory_listing: bool,
}

/// Lazy cache of per-directory `AspDirConfig` resolved from `asp.ini` files.
///
/// Resolution walks from the root folder's immediate children down to the
/// requested directory, loading each `asp.ini`'s `[server]` section in order.
/// Deeper INI files override keys set by shallower ones (closest wins).
/// The base defaults apply when no INI is present along the path.
/// Subsequent lookups are O(1) cache hits.
pub struct DirConfigCache {
    base: AspDirConfig,
    root_folder: PathBuf,
    cache: RwLock<HashMap<PathBuf, AspDirConfig>>,
}

impl DirConfigCache {
    /// Create a new cache with the given base defaults and canonical root folder.
    pub fn new(base: AspDirConfig, root_folder: PathBuf) -> Self {
        DirConfigCache {
            base,
            root_folder,
            cache: RwLock::new(HashMap::new()),
        }
    }

    /// Resolve the effective `AspDirConfig` for a canonical directory path.
    ///
    /// Algorithm:
    /// 1. Check the write-locked cache — return clone on hit.
    /// 2. Start from the base config.
    /// 3. Walk each path component under root_folder; if `asp.ini` exists,
    ///    merge its `[server]` section on top of the running config.
    /// 4. Store the result in cache and return it.
    pub fn resolve(&self, canonical_dir: &Path) -> AspDirConfig {
        if let Some(config) = self.cache.read().unwrap().get(canonical_dir) {
            return config.clone();
        }

        let mut dir_config = self.base.clone();

        if let Ok(relative) = canonical_dir.strip_prefix(&self.root_folder) {
            let mut current = self.root_folder.clone();
            for component in relative.components() {
                current = current.join(component);
                let ini_path = current.join("asp.ini");
                if ini_path.is_file() {
                    if let Ok(content) = std::fs::read_to_string(&ini_path) {
                        Self::apply_ini_to_dir_config(&mut dir_config, &content);
                    }
                }
            }
        }

        self.cache
            .write()
            .unwrap()
            .insert(canonical_dir.to_path_buf(), dir_config.clone());
        dir_config
    }

    /// Parse the `[server]` section of an `asp.ini` file and merge its
    /// key-value pairs into `dir_config`. Ignores comments (`#`, `;`),
    /// empty lines, and sections other than `[server]`.
    ///
    /// Supported keys:
    /// - `default_documents` — comma-separated list (replaces the whole list)
    /// - `default_document`  — single value (backward compat, replaces the list)
    /// - `enable_directory_listing` — boolean (`true` enables)
    fn apply_ini_to_dir_config(dir_config: &mut AspDirConfig, content: &str) {
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
                    "default_documents" => {
                        dir_config.default_documents = value
                            .split(',')
                            .map(|s| s.trim().to_string())
                            .filter(|s| !s.is_empty())
                            .collect();
                    }
                    "default_document" => {
                        if !value.is_empty() {
                            dir_config.default_documents = vec![value.to_string()];
                        }
                    }
                    "enable_directory_listing" => {
                        dir_config.directory_listing = value.eq_ignore_ascii_case("true");
                    }
                    _ => {}
                }
            }
        }
    }
}

/// Runtime server configuration with all overrides applied in order:
/// defaults < INI file < programmatic overrides.
#[derive(Debug, Clone)]
pub struct AspServerConfig {
    pub host: String,
    pub port: u16,
    pub folder: String,
    /// Prioritized list of default documents (IIS-like fallback chain).
    pub default_documents: Vec<String>,
    pub directory_listing: bool,
}

impl Default for AspServerConfig {
    fn default() -> Self {
        Self {
            host: "127.0.0.1".to_string(),
            port: 9090,
            folder: ".".to_string(),
            default_documents: vec![
                "default.asp".into(),
                "default.htm".into(),
                "default.html".into(),
                "index.asp".into(),
                "index.htm".into(),
                "index.html".into(),
                "iisstart.htm".into(),
            ],
            directory_listing: false,
        }
    }
}

impl AspServerConfig {
    /// Load `asp.ini` from the served folder and apply its values on top of defaults.
    ///
    /// Reads the `[server]` section of `<folder>/asp.ini` and applies recognized keys.
    /// This is the per-server-root INI; per-directory INI files are handled by
    /// `DirConfigCache` at request time.
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
                        "default_documents" => {
                            cfg.default_documents = value
                                .split(',')
                                .map(|s| s.trim().to_string())
                                .filter(|s| !s.is_empty())
                                .collect();
                        }
                        "default_document" => {
                            if !value.is_empty() {
                                cfg.default_documents = vec![value.to_string()];
                            }
                        }
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

    /// Build a `DirConfigCache` from this server config's values.
    pub fn build_dir_cache(&self) -> DirConfigCache {
        let root = Path::new(&self.folder)
            .canonicalize()
            .unwrap_or_else(|_| Path::new(&self.folder).to_path_buf());
        DirConfigCache::new(
            AspDirConfig {
                default_documents: self.default_documents.clone(),
                directory_listing: self.directory_listing,
            },
            root,
        )
    }

    /// Apply overrides from external sources (e.g. DAP launch args or CLI args).
    ///
    /// Override priority (highest wins):
    ///   defaults < asp.ini < `apply_overrides`
    /// Empty / `None` values are skipped so INI/defaults are preserved.
    /// `default_documents` is a comma-separated string that replaces the full list.
    pub fn apply_overrides(&mut self, host: Option<&str>, port: Option<u16>, folder: Option<&str>, default_documents: Option<&str>, directory_listing: Option<bool>) {
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
        if let Some(d) = default_documents {
            if !d.is_empty() {
                self.default_documents = d
                    .split(',')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
                    .collect();
            }
        }
        if let Some(dl) = directory_listing {
            self.directory_listing = dl;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_asp_server_config_defaults_are_iis_like() {
        let cfg = AspServerConfig::default();
        assert!(cfg.default_documents.len() >= 7);
        assert!(cfg.default_documents.contains(&"default.asp".to_string()));
        assert!(cfg.default_documents.contains(&"index.asp".to_string()));
        assert!(cfg.default_documents.contains(&"iisstart.htm".to_string()));
        assert!(!cfg.directory_listing);
    }

    #[test]
    fn test_asp_server_config_from_folder_loads_ini() {
        let dir = std::env::temp_dir().join(format!("asp_test_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir);
        std::fs::write(
            dir.join("asp.ini"),
            "[server]\ndefault_documents = foo.asp, bar.htm\nenable_directory_listing = true\n",
        )
        .unwrap();
        let cfg = AspServerConfig::from_folder(dir.to_str().unwrap());
        assert_eq!(cfg.default_documents, vec!["foo.asp", "bar.htm"]);
        assert!(cfg.directory_listing);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_asp_server_config_from_folder_default_document_compat() {
        let dir = std::env::temp_dir().join(format!("asp_test_single_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir);
        std::fs::write(dir.join("asp.ini"), "[server]\ndefault_document = legacy.asp\n").unwrap();
        let cfg = AspServerConfig::from_folder(dir.to_str().unwrap());
        assert_eq!(cfg.default_documents, vec!["legacy.asp"]);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_asp_server_config_apply_overrides_replaces_list() {
        let mut cfg = AspServerConfig::default();
        cfg.apply_overrides(None, None, None, Some("override.asp, fallback.htm"), None);
        assert_eq!(cfg.default_documents, vec!["override.asp", "fallback.htm"]);
    }

    #[test]
    fn test_asp_server_config_apply_overrides_empty_ignored() {
        let mut cfg = AspServerConfig::default();
        let orig = cfg.default_documents.clone();
        cfg.apply_overrides(None, None, None, Some(""), None);
        assert_eq!(cfg.default_documents, orig);
    }

    #[test]
    fn test_dir_config_cache_resolve_root() {
        let dir = std::env::temp_dir().join(format!("asp_cache_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir);
        let base = AspDirConfig {
            default_documents: vec!["index.asp".to_string()],
            directory_listing: false,
        };
        let root = dir.canonicalize().unwrap();
        let cache = DirConfigCache::new(base.clone(), root.clone());

        let resolved = cache.resolve(&root);
        assert_eq!(resolved.default_documents, vec!["index.asp"]);
        assert!(!resolved.directory_listing);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_dir_config_cache_resolve_subdir_with_ini() {
        let dir = std::env::temp_dir().join(format!("asp_cache_sub_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir.join("sub"));
        std::fs::write(
            dir.join("sub").join("asp.ini"),
            "[server]\ndefault_documents = subpage.asp\n",
        )
        .unwrap();

        let base = AspDirConfig {
            default_documents: vec!["index.asp".to_string()],
            directory_listing: false,
        };
        let root = dir.canonicalize().unwrap();
        let sub = root.join("sub");
        let cache = DirConfigCache::new(base, root);

        let resolved = cache.resolve(&sub);
        assert_eq!(resolved.default_documents, vec!["subpage.asp"]);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_dir_config_cache_resolve_nested_merge() {
        let dir = std::env::temp_dir().join(format!("asp_cache_nest_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir.join("sub").join("deep"));
        std::fs::write(
            dir.join("sub").join("asp.ini"),
            "[server]\nenable_directory_listing = true\n",
        )
        .unwrap();
        std::fs::write(
            dir.join("sub").join("deep").join("asp.ini"),
            "[server]\ndefault_documents = deep.asp\n",
        )
        .unwrap();

        let base = AspDirConfig {
            default_documents: vec!["index.asp".to_string()],
            directory_listing: false,
        };
        let root = dir.canonicalize().unwrap();
        let deep = root.join("sub").join("deep");
        let cache = DirConfigCache::new(base, root);

        let resolved = cache.resolve(&deep);
        assert_eq!(resolved.default_documents, vec!["deep.asp"]);
        assert!(resolved.directory_listing); // inherited from sub/asp.ini
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_dir_config_cache_caches_results() {
        let dir = std::env::temp_dir().join(format!("asp_cache_hit_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir);
        let base = AspDirConfig {
            default_documents: vec!["a.asp".to_string()],
            directory_listing: false,
        };
        let root = dir.canonicalize().unwrap();
        let cache = DirConfigCache::new(base, root.clone());

        let r1 = cache.resolve(&root);
        let r2 = cache.resolve(&root);
        assert_eq!(r1.default_documents, r2.default_documents);
        // Second call should not create filesystem side effects — test passes if no panic
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_dir_config_cache_inherits_base_without_ini() {
        let dir = std::env::temp_dir().join(format!("asp_cache_inherit_{}", std::process::id()));
        let _ = std::fs::create_dir_all(&dir.join("empty"));
        let base = AspDirConfig {
            default_documents: vec!["base.asp".to_string()],
            directory_listing: true,
        };
        let root = dir.canonicalize().unwrap();
        let empty = root.join("empty");
        let cache = DirConfigCache::new(base, root);

        let resolved = cache.resolve(&empty);
        assert_eq!(resolved.default_documents, vec!["base.asp"]);
        assert!(resolved.directory_listing);
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn test_asp_dir_config_clone() {
        let a = AspDirConfig {
            default_documents: vec!["x.asp".to_string()],
            directory_listing: true,
        };
        let b = a.clone();
        assert_eq!(a.default_documents, b.default_documents);
        assert_eq!(a.directory_listing, b.directory_listing);
    }
}
