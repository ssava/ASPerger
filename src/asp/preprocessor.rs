//! Preprocessor for `<%@ ... %>` ASP directives. Filters directive blocks
//! from the parsed block list and returns a `DirectiveConfig` with settings
//! for language, session state, code page, LCID, and transaction.

use crate::asp::parser::AspBlock;

/// Configuration extracted from `<%@ ... %>` directives.
#[derive(Debug, Clone)]
pub struct DirectiveConfig {
    pub language: String,
    pub enable_session_state: bool,
    pub transaction: Option<String>,
    pub code_page: Option<u32>,
    pub lcid: Option<u32>,
}

impl Default for DirectiveConfig {
    fn default() -> Self {
        DirectiveConfig {
            language: "VBScript".to_string(),
            enable_session_state: true,
            transaction: None,
            code_page: None,
            lcid: None,
        }
    }
}

/// Processes `AspBlock::Directive` blocks extracted by the parser.
pub struct Preprocessor;

impl Preprocessor {
    /// Create a new Preprocessor.
    pub fn new() -> Self {
        Preprocessor
    }

    /// Scan blocks for AspBlock::Directive, accumulate into DirectiveConfig,
    /// and return non-directive block references.
    pub fn process<'a>(&self, blocks: &'a [AspBlock]) -> (DirectiveConfig, Vec<&'a AspBlock>) {
        let mut config = DirectiveConfig::default();
        let mut filtered = Vec::new();

        for block in blocks {
            match block {
                AspBlock::Directive(name, value) => {
                    match name.to_uppercase().as_str() {
                        "LANGUAGE" => config.language = value.clone(),
                        "ENABLESESSIONSTATE" => {
                            config.enable_session_state =
                                value.eq_ignore_ascii_case("true")
                        }
                        "TRANSACTION" => config.transaction = Some(value.clone()),
                        "CODEPAGE" => {
                            config.code_page = value.parse().ok();
                        }
                        "LCID" => {
                            config.lcid = value.parse().ok();
                        }
                        _ => {}
                    }
                }
                other => filtered.push(other),
            }
        }

        (config, filtered)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::asp::parser::AspBlock;

    #[test]
    fn test_process_no_directives() {
        let blocks = vec![
            AspBlock::Html("<html>".to_string()),
            AspBlock::Code("x = 1".to_string()),
        ];
        let p = Preprocessor::new();
        let (config, filtered) = p.process(&blocks);
        assert_eq!(config.language, "VBScript");
        assert!(config.enable_session_state);
        assert_eq!(filtered.len(), 2);
    }

    #[test]
    fn test_process_directives() {
        let blocks = vec![
            AspBlock::Directive("LANGUAGE".to_string(), "VBScript".to_string()),
            AspBlock::Directive(
                "ENABLESESSIONSTATE".to_string(),
                "False".to_string(),
            ),
            AspBlock::Directive("CODEPAGE".to_string(), "65001".to_string()),
            AspBlock::Html("content".to_string()),
        ];
        let p = Preprocessor::new();
        let (config, filtered) = p.process(&blocks);
        assert_eq!(config.language, "VBScript");
        assert!(!config.enable_session_state);
        assert_eq!(config.code_page, Some(65001));
        assert_eq!(filtered.len(), 1);
        match filtered[0] {
            AspBlock::Html(s) => assert_eq!(s, "content"),
            _ => panic!("Expected Html block"),
        }
    }

    #[test]
    fn test_process_lcid_transaction() {
        let blocks = vec![
            AspBlock::Directive("LCID".to_string(), "1033".to_string()),
            AspBlock::Directive(
                "TRANSACTION".to_string(),
                "Required".to_string(),
            ),
        ];
        let p = Preprocessor::new();
        let (config, _) = p.process(&blocks);
        assert_eq!(config.lcid, Some(1033));
        assert_eq!(config.transaction, Some("Required".to_string()));
    }

    #[test]
    fn test_process_unknown_directive_ignored() {
        let blocks = vec![
            AspBlock::Directive("UNKNOWN".to_string(), "value".to_string()),
        ];
        let p = Preprocessor::new();
        let (config, filtered) = p.process(&blocks);
        assert!(filtered.is_empty());
        assert_eq!(config.language, "VBScript");
    }
}
