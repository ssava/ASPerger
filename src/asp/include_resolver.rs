use regex::Regex;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

const MAX_INCLUDE_DEPTH: usize = 10;

fn get_include_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(r#"<!--\s*#include\s+(file|virtual)\s*=\s*"([^"]+)"\s*-->"#).unwrap()
    })
}

pub struct IncludeResolver;

impl IncludeResolver {
    pub fn expand(
        source: &str,
        base_dir: &Path,
        root_dir: &Path,
    ) -> Result<String, String> {
        let mut path_stack = Vec::new();
        Self::expand_recursive(source, base_dir, root_dir, &mut path_stack, 0)
    }

    fn expand_recursive(
        source: &str,
        base_dir: &Path,
        root_dir: &Path,
        path_stack: &mut Vec<PathBuf>,
        depth: usize,
    ) -> Result<String, String> {
        if depth > MAX_INCLUDE_DEPTH {
            return Err(format!(
                "Maximum include depth ({}) exceeded",
                MAX_INCLUDE_DEPTH
            ));
        }

        let re = get_include_regex();
        let mut result = String::new();
        let mut last_end = 0;

        for cap in re.captures_iter(source) {
            let m = cap.get(0).unwrap();
            result.push_str(&source[last_end..m.start()]);

            let include_type = cap.get(1).unwrap().as_str();
            let include_path = cap.get(2).unwrap().as_str();

            let resolved = if include_type == "file" {
                base_dir.join(include_path)
            } else {
                let trimmed = include_path.trim_start_matches('/');
                root_dir.join(trimmed)
            };

            let canonical = resolved.canonicalize().map_err(|e| {
                format!("Include file not found '{}': {}", resolved.display(), e)
            })?;

            if path_stack.contains(&canonical) {
                return Err(format!(
                    "Circular include detected: {}",
                    canonical.display()
                ));
            }

            let included = std::fs::read_to_string(&canonical).map_err(|e| {
                format!(
                    "Could not read include '{}': {}",
                    canonical.display(),
                    e
                )
            })?;

            path_stack.push(canonical.clone());
            let expanded = Self::expand_recursive(
                &included,
                canonical.parent().unwrap_or(base_dir),
                root_dir,
                path_stack,
                depth + 1,
            )?;
            path_stack.pop();

            result.push_str(&expanded);
            last_end = m.end();
        }

        result.push_str(&source[last_end..]);
        Ok(result)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    #[test]
    fn test_expand_no_includes() {
        let source = "<html><body><%= \"Hello\" %></body></html>";
        let dir = std::env::temp_dir();
        let result = IncludeResolver::expand(source, &dir, &dir).unwrap();
        assert_eq!(result, source);
    }

    #[test]
    fn test_expand_single_include() {
        let dir = std::env::temp_dir().join("include_test_single");
        let _ = fs::create_dir_all(&dir);
        let inc_path = dir.join("header.inc");
        fs::write(&inc_path, "<h1>Header</h1>").unwrap();

        let source = "before<!-- #include file=\"header.inc\" -->after";
        let result =
            IncludeResolver::expand(source, &dir, &dir).unwrap();

        assert_eq!(result, "before<h1>Header</h1>after");
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_expand_nested_includes() {
        let dir = std::env::temp_dir().join("include_test_nested");
        let _ = fs::create_dir_all(&dir);
        fs::write(dir.join("a.inc"), "A<!-- #include file=\"b.inc\" -->A").unwrap();
        fs::write(dir.join("b.inc"), "B<!-- #include file=\"c.inc\" -->B").unwrap();
        fs::write(dir.join("c.inc"), "C").unwrap();

        let result = IncludeResolver::expand("X<!-- #include file=\"a.inc\" -->X", &dir, &dir).unwrap();
        assert_eq!(result, "XABCBAX");
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_expand_circular_include_error() {
        let dir = std::env::temp_dir().join("include_test_circular");
        let _ = fs::create_dir_all(&dir);
        fs::write(dir.join("a.inc"), "A<!-- #include file=\"b.inc\" -->").unwrap();
        fs::write(dir.join("b.inc"), "B<!-- #include file=\"a.inc\" -->").unwrap();

        let result = IncludeResolver::expand("X<!-- #include file=\"a.inc\" -->X", &dir, &dir);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Circular"));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_expand_virtual_include() {
        let dir = std::env::temp_dir().join("include_test_virtual");
        let sub = dir.join("sub");
        let _ = fs::create_dir_all(&sub);
        fs::write(sub.join("footer.inc"), "FOOTER").unwrap();

        let source = "body<!-- #include virtual=\"/sub/footer.inc\" -->end";
        let result = IncludeResolver::expand(source, &sub, &dir).unwrap();
        assert_eq!(result, "bodyFOOTERend");
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_expand_missing_file_error() {
        let dir = std::env::temp_dir().join("include_test_missing");
        let _ = fs::create_dir_all(&dir);

        let result = IncludeResolver::expand(
            "<!-- #include file=\"nope.inc\" -->",
            &dir,
            &dir,
        );
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("not found"));
        fs::remove_dir_all(&dir).ok();
    }
}
