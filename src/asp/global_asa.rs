//! Global.asa file parser. Handles the special ASP application file
//! that defines Application_OnStart, Application_OnEnd, Session_OnStart,
//! Session_OnEnd event handlers and <OBJECT> declarations.

use regex::Regex;
use std::sync::OnceLock;

/// A single `<OBJECT>` declaration from global.asa.
#[derive(Debug, Clone)]
pub struct ObjectDecl {
    pub id: String,
    pub progid: String,
    pub scope: ObjectScope,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ObjectScope {
    Application,
    Session,
}

/// Parsed contents of a global.asa file.
#[derive(Debug, Default, Clone)]
pub struct GlobalAsa {
    /// Raw VBScript source of `Sub Application_OnStart ... End Sub`
    pub app_on_start: Option<String>,
    /// Raw VBScript source of `Sub Application_OnEnd ... End Sub`
    pub app_on_end: Option<String>,
    /// Raw VBScript source of `Sub Session_OnStart ... End Sub`
    pub session_on_start: Option<String>,
    /// Raw VBScript source of `Sub Session_OnEnd ... End Sub`
    pub session_on_end: Option<String>,
    /// Application-scoped `<OBJECT>` declarations.
    pub app_objects: Vec<ObjectDecl>,
    /// Session-scoped `<OBJECT>` declarations.
    pub session_objects: Vec<ObjectDecl>,
    /// Any module-level VBScript code (non-event-handler) from global.asa.
    /// Executed once during Application_OnStart.
    pub module_code: Vec<String>,
}

fn get_script_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        // Match any <SCRIPT ...>...</SCRIPT>  (lookahead not supported by regex crate)
        Regex::new(r"(?is)<\s*SCRIPT\b([^>]*)>(.*?)<\s*/\s*SCRIPT\s*>").unwrap()
    })
}

fn get_object_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        // Match any <OBJECT ...>...</OBJECT>
        Regex::new(r"(?is)<\s*OBJECT\b([^>]*)>(.*?)<\s*/\s*OBJECT\s*>").unwrap()
    })
}

fn get_object_attr_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        // Match SCOPE/ID/PROGID="value" or SCOPE/ID/PROGID=value
        Regex::new(
            r#"(?is)\b(SCOPE|ID|PROGID)\s*=\s*"?\s*([^"'\s>]+)\s*"?"#,
        )
        .unwrap()
    })
}

/// Parse the content of a global.asa file.
pub fn parse_global_asa(content: &str) -> GlobalAsa {
    let mut result = GlobalAsa::default();

    // Step 1: Extract <OBJECT> tags, remove from content
    let content = get_object_regex().replace_all(content, |caps: &regex::Captures| {
        let attrs = caps.get(1).map_or("", |m| m.as_str());
        // Only process if RUNAT="Server" is present
        if !attrs.to_uppercase().contains("SERVER") {
            return caps.get(0).map_or(String::new(), |m| m.as_str().to_string());
        }
        let full_tag = caps.get(0).map_or("", |m| m.as_str());
        let attr_re = get_object_attr_regex();
        let mut id = String::new();
        let mut progid = String::new();
        let mut scope_str = String::new();
        for attr_cap in attr_re.captures_iter(full_tag) {
            let name = attr_cap.get(1).map_or("", |m| m.as_str()).to_uppercase();
            let value = attr_cap.get(2).map_or("", |m| m.as_str());
            match name.as_str() {
                "ID" => id = value.to_string(),
                "PROGID" => progid = value.to_string(),
                "SCOPE" => scope_str = value.to_string(),
                _ => {}
            }
        }
        if !id.is_empty() && !progid.is_empty() {
            let decl = ObjectDecl {
                id,
                progid,
                scope: if scope_str.eq_ignore_ascii_case("Application") {
                    ObjectScope::Application
                } else {
                    ObjectScope::Session
                },
            };
            match decl.scope {
                ObjectScope::Application => result.app_objects.push(decl),
                ObjectScope::Session => result.session_objects.push(decl),
            }
        }
        String::new()
    }).to_string();

    // Step 2: Convert <SCRIPT RUNAT="Server"> blocks to <% %> blocks
    let content = get_script_regex()
        .replace_all(&content, |caps: &regex::Captures| {
            let attrs = caps.get(1).map_or("", |m| m.as_str());
            // Only convert if LANGUAGE=VBScript and RUNAT=Server
            let upper_attrs = attrs.to_uppercase();
            if !upper_attrs.contains("VBSCRIPT") || !upper_attrs.contains("SERVER") {
                return caps.get(0).map_or(String::new(), |m| m.as_str().to_string());
            }
            let inner = caps.get(2).map_or("", |m| m.as_str());
            format!("<%{}%>", inner.trim())
        })
        .to_string();

    // Step 3: Parse with the standard AspParser
    let parser = crate::asp::parser::AspParser::new(content);
    let blocks = parser.parse();

    // Step 4: Scan code blocks for event handler Sub definitions
    for block in &blocks {
        if let crate::asp::parser::AspBlock::Code(code, _) = block {
            let code_upper = code.to_uppercase();

            // Collect (handler_name, byte_range_start, extracted_code) for each found handler
            let mut handler_ranges: Vec<(usize, usize)> = Vec::new();

            let handler_names = [
                "APPLICATION_ONSTART",
                "APPLICATION_ONEND",
                "SESSION_ONSTART",
                "SESSION_ONEND",
            ];

            for handler_name in &handler_names {
                if !code_upper.contains(handler_name) {
                    continue;
                }
                let sub_keyword = "SUB ";
                let mut search_start = 0;
                while let Some(sub_pos) = code_upper[search_start..].find(sub_keyword) {
                    let abs_pos = search_start + sub_pos;
                    let after_sub = &code_upper[abs_pos + sub_keyword.len()..];
                    if after_sub.trim_start().starts_with(handler_name) {
                        let code_after = &code[abs_pos..];
                        let code_after_upper = &code_upper[abs_pos..];
                        if let Some(end_sub_pos) = code_after_upper.find("END SUB") {
                            let end_abs = abs_pos + end_sub_pos + "END SUB".len();
                            let handler_code = code_after[..end_sub_pos + "END SUB".len()].trim().to_string();
                            // Store in the appropriate slot
                            let slot: &mut Option<String> = match *handler_name {
                                "APPLICATION_ONSTART" => &mut result.app_on_start,
                                "APPLICATION_ONEND" => &mut result.app_on_end,
                                "SESSION_ONSTART" => &mut result.session_on_start,
                                "SESSION_ONEND" => &mut result.session_on_end,
                                _ => unreachable!(),
                            };
                            if slot.is_none() {
                                *slot = Some(handler_code);
                                handler_ranges.push((abs_pos, end_abs));
                            }
                            break;
                        }
                    }
                    search_start = abs_pos + 1;
                }
            }

            // Compute module-level code: everything not in a handler Sub...End Sub range
            if !handler_ranges.is_empty() {
                handler_ranges.sort_unstable();
                let mut remaining = String::new();
                let mut last_end = 0;
                for (start, end) in &handler_ranges {
                    remaining.push_str(&code[last_end..*start]);
                    last_end = *end;
                }
                remaining.push_str(&code[last_end..]);
                let trimmed = remaining.trim();
                if !trimmed.is_empty() {
                    result.module_code.push(trimmed.to_string());
                }
            } else {
                result.module_code.push(code.clone());
            }
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_empty() {
        let g = parse_global_asa("");
        assert!(g.app_on_start.is_none());
        assert!(g.session_on_start.is_none());
        assert!(g.app_objects.is_empty());
        assert!(g.session_objects.is_empty());
    }

    #[test]
    fn test_parse_script_syntax() {
        let content = r#"
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub Application_OnStart
    Application("counter") = 0
End Sub
</SCRIPT>
"#;
        let g = parse_global_asa(content);
        assert!(g.app_on_start.is_some());
        let code = g.app_on_start.unwrap();
        assert!(code.contains("Application_OnStart"));
        assert!(code.contains("counter"));
        assert!(code.contains("End Sub"));
    }

    #[test]
    fn test_parse_percent_syntax() {
        let content = r#"
<%
Sub Application_OnStart
    Application("counter") = 0
End Sub
%>
"#;
        let g = parse_global_asa(content);
        assert!(g.app_on_start.is_some());
    }

    #[test]
    fn test_parse_object_tag() {
        let content = r#"
<OBJECT RUNAT="Server" SCOPE="Application" ID="MyDict" PROGID="Scripting.Dictionary"></OBJECT>
<OBJECT RUNAT="Server" SCOPE="Session" ID="MyConn" PROGID="ADODB.Connection"></OBJECT>
"#;
        let g = parse_global_asa(content);
        assert_eq!(g.app_objects.len(), 1);
        assert_eq!(g.app_objects[0].id, "MyDict");
        assert_eq!(g.app_objects[0].progid, "Scripting.Dictionary");
        assert_eq!(g.app_objects[0].scope, ObjectScope::Application);
        assert_eq!(g.session_objects.len(), 1);
        assert_eq!(g.session_objects[0].id, "MyConn");
        assert_eq!(g.session_objects[0].progid, "ADODB.Connection");
        assert_eq!(g.session_objects[0].scope, ObjectScope::Session);
    }

    #[test]
    fn test_parse_all_events() {
        let content = r#"
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub Application_OnStart
    Application("started") = Now()
End Sub

Sub Application_OnEnd
    ' cleanup
End Sub

Sub Session_OnStart
    Session("visited") = 1
End Sub

Sub Session_OnEnd
    ' session cleanup
End Sub
</SCRIPT>
"#;
        let g = parse_global_asa(content);
        assert!(g.app_on_start.is_some());
        assert!(g.app_on_end.is_some());
        assert!(g.session_on_start.is_some());
        assert!(g.session_on_end.is_some());

        assert!(g.app_on_start.unwrap().contains("Now()"));
        assert!(g.session_on_start.unwrap().contains("visited"));
    }

    #[test]
    fn test_parse_module_code() {
        let content = r#"
<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Dim g_counter
g_counter = 0

Sub Application_OnStart
    Application("started") = Now()
End Sub
</SCRIPT>
"#;
        let g = parse_global_asa(content);
        assert!(g.app_on_start.is_some());
        // The module-level `Dim g_counter` and `g_counter = 0` are separate blocks
        assert!(!g.module_code.is_empty());
    }

    #[test]
    fn test_parse_mixed_syntax() {
        let content = r#"
<OBJECT RUNAT="Server" SCOPE="Application" ID="AppData" PROGID="Scripting.Dictionary"></OBJECT>

<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
Sub Application_OnStart
    AppData.Add "key", "value"
End Sub
</SCRIPT>

<% Session("theme") = "dark" %>
"#;
        let g = parse_global_asa(content);
        assert_eq!(g.app_objects.len(), 1);
        assert!(g.app_on_start.is_some());
    }
}
