//! ASP block parser. Splits raw ASP source text into `AspBlock` variants
//! (Html, Code, Directive), handling `<%= %>` expression shorthand and
//! `<%@ %>` directive syntax via regex.

use regex::Regex;
use std::sync::OnceLock;

/// Represents a single parsed segment of an ASP page.
#[derive(Debug)]
pub enum AspBlock {
    /// Literal HTML content to be sent to the client.
    Html(String),
    /// VBScript code between `<% %>` delimiters, with the ASP file line number
    /// where this block's first code character appears.
    Code(String, usize),
    /// An `<%@ ... %>` directive with name and value.
    Directive(String, String),
}

/// Parses ASP source text into a sequence of `AspBlock` values.
pub struct AspParser {
    content: String,
}

fn get_asp_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r"<%(?s)(.*?)%>").unwrap())
}

fn get_asp_expression_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r"<%=\s*(.*?)\s*%>").unwrap())
}

fn get_asp_directive_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r#"<%@\s*(\w+)\s*=\s*(?:"([^"]*)"|(\w+))\s*%>"#).unwrap())
}

impl AspParser {
    pub fn new(content: String) -> Self {
        AspParser { content }
    }

    pub fn parse(&self) -> Vec<AspBlock> {
        let mut blocks = Vec::new();
        let expr_re = get_asp_expression_regex();
        let dir_re = get_asp_directive_regex();
        let re = get_asp_regex();
        let mut last_end = 0;

        // Pre-scan original content for Code block line numbers.
        // This happens before transformations so we get accurate ASP file line numbers.
        // We skip directives and empty blocks, matching the downstream filter.
        let code_block_lines: Vec<usize> = re
            .captures_iter(&self.content)
            .filter_map(|cap| {
                let code = cap.get(1).map_or("", |m| m.as_str());
                if code.trim().is_empty() || code.trim_start().starts_with('@') {
                    return None;
                }
                let m = cap.get(0).unwrap();
                let line = self.content[..m.start()].matches('\n').count() + 1;
                if code.starts_with('\n') {
                    Some(line + 1)
                } else {
                    Some(line)
                }
            })
            .collect();

        // First pass: handle <%= expr %> shorthand
        let content = expr_re.replace_all(&self.content, |caps: &regex::Captures| {
            let expr = caps.get(1).map_or("", |m| m.as_str());
            format!("<% Response.Write({}) %>", expr)
        });

        // Second pass: strip directives and add them as Directive blocks
        let content = dir_re
            .replace_all(&content, |caps: &regex::Captures| {
                let name = caps.get(1).map_or("", |m| m.as_str());
                let value = caps.get(2).map_or("", |m| m.as_str()).to_owned();
                let value = if value.is_empty() {
                    caps.get(3).map_or("", |m| m.as_str()).to_string()
                } else {
                    value
                };
                blocks.push(AspBlock::Directive(name.to_string(), value));
                ""
            })
            .to_string();

        let mut code_idx = 0;

        for cap in re.captures_iter(&content) {
            let whole_match = cap.get(0).unwrap();
            let code = cap.get(1).map_or("", |m| m.as_str());

            if whole_match.start() > last_end {
                let html = &content[last_end..whole_match.start()];
                blocks.push(AspBlock::Html(html.to_string()));
            }

            if !code.trim().is_empty() {
                let line = *code_block_lines.get(code_idx).unwrap_or(&1);
                blocks.push(AspBlock::Code(code.trim().to_string(), line));
                code_idx += 1;
            }

            last_end = whole_match.end();
        }

        if last_end < content.len() {
            let html = &content[last_end..];
            blocks.push(AspBlock::Html(html.to_string()));
        }

        tracing::trace!(count = blocks.len(), "Parsed ASP blocks");
        blocks
    }
}
