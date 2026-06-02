use regex::Regex;
use std::sync::OnceLock;

#[derive(Debug)]
pub enum AspBlock {
    Html(String),
    Code(String),
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

impl AspParser {
    pub fn new(content: String) -> Self {
        AspParser { content }
    }

    pub fn parse(&self) -> Vec<AspBlock> {
        let mut blocks = Vec::new();
        let expr_re = get_asp_expression_regex();
        let re = get_asp_regex();
        let mut last_end = 0;

        // First pass: handle <%= expr %> shorthand
        let content = expr_re.replace_all(&self.content, |caps: &regex::Captures| {
            let expr = caps.get(1).map_or("", |m| m.as_str());
            format!("<% Response.Write({}) %>", expr)
        });

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