use regex::Regex;

#[derive(Debug)]
pub enum AspBlock {
    Html(String),
    Code(String),
}

pub struct AspParser {
    content: String,
}

impl AspParser {
    pub fn new(content: String) -> Self {
        AspParser { content }
    }

    pub fn parse(&self) -> Vec<AspBlock> {
        let mut blocks = Vec::new();
        let re = Regex::new(r"<%(?s)(.*?)%>").unwrap();
        let mut last_end = 0;

        for cap in re.captures_iter(&self.content) {
            let whole_match = cap.get(0).unwrap();
            let code = cap.get(1).map_or("", |m| m.as_str());
            
            if whole_match.start() > last_end {
                let html = &self.content[last_end..whole_match.start()];
                if !html.trim().is_empty() {
                    blocks.push(AspBlock::Html(html.to_string()));
                }
            }

            if !code.trim().is_empty() {
                blocks.push(AspBlock::Code(code.trim().to_string()));
            }
            
            last_end = whole_match.end();
        }

        if last_end < self.content.len() {
            let html = &self.content[last_end..];
            if !html.trim().is_empty() {
                blocks.push(AspBlock::Html(html.to_string()));
            }
        }

        blocks
    }
}