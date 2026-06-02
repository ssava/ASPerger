use regex::Regex;
use std::sync::OnceLock;

#[derive(Debug)]
pub enum AspBlock {
    Html(String),
    Code(String),
    Directive(String, String),
}

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
    RE.get_or_init(|| Regex::new(r"<%@\s*(\w+)\s*=\s*(\w+)\s*%>").unwrap())
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

        // First pass: handle <%= expr %> shorthand
        let content = expr_re.replace_all(&self.content, |caps: &regex::Captures| {
            let expr = caps.get(1).map_or("", |m| m.as_str());
            format!("<% Response.Write({}) %>", expr)
        });

        // Second pass: strip directives and add them as Directive blocks
        let content = dir_re.replace_all(&content, |caps: &regex::Captures| {
            let name = caps.get(1).map_or("", |m| m.as_str());
            let value = caps.get(2).map_or("", |m| m.as_str());
            blocks.push(AspBlock::Directive(name.to_string(), value.to_string()));
            ""
        }).to_string();

        for cap in re.captures_iter(&content) {
            let whole_match = cap.get(0).unwrap();
            let code = cap.get(1).map_or("", |m| m.as_str());

            if whole_match.start() > last_end {
                let html = &content[last_end..whole_match.start()];
                blocks.push(AspBlock::Html(html.to_string()));
            }

            if !code.trim().is_empty() {
                blocks.push(AspBlock::Code(code.trim().to_string()));
            }

            last_end = whole_match.end();
        }

        if last_end < content.len() {
            let html = &content[last_end..];
            blocks.push(AspBlock::Html(html.to_string()));
        }

        blocks
    }
}