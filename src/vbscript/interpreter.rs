use crate::asp::parser::AspBlock;
use crate::vbscript::block;
use crate::vbscript::vbobject::ErrObject;
use crate::vbscript::vbs_error::VBSError;
use crate::vbscript::ExecutionContext;
use crate::vbscript::{Token, TokenType, Tokenizer, VBValue};

pub struct VBScriptInterpreter;

impl VBScriptInterpreter {
    pub fn execute(&self, code: &str, context: &mut ExecutionContext) -> Result<(), VBSError> {
        self.execute_vm(code, context)
    }

    pub fn execute_vm(&self, code: &str, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let code = code.trim().to_string();

        let tokens = Tokenizer::tokenize(&code);
        if tokens.iter().all(|t| t.token_type == TokenType::EOF) {
            return Ok(());
        }

        tracing::trace!(token_count = tokens.len(), "Tokenized code block");

        let lines = self.group_tokens_into_lines(&tokens)?;

        if context.get_variable("ERR").is_none() {
            context.set_variable("ERR", VBValue::Object(Box::new(ErrObject::new())));
        }

        inject_vbscript_constants(context);

        let blocks = block::parse_blocks(&lines)?;
        tracing::trace!(block_count = blocks.len(), "Parsed VBScript blocks");

        self.run_compiled_blocks(&blocks, context)
    }

    /// Execute pre-parsed `BlockStatement`s via the VM.
    pub fn execute_blocks_vm(&self, blocks: &[block::BlockStatement], context: &mut ExecutionContext) -> Result<(), VBSError> {
        if context.get_variable("ERR").is_none() {
            context.set_variable("ERR", VBValue::Object(Box::new(ErrObject::new())));
        }

        inject_vbscript_constants(context);

        self.run_compiled_blocks(blocks, context)
    }

    fn run_compiled_blocks(&self, blocks: &[block::BlockStatement], context: &mut ExecutionContext) -> Result<(), VBSError> {
        let mut compiler = crate::vbscript::compiler::Compiler::new(context);
        let mut compiled = compiler.compile(blocks)?;

        for func in compiled.function_defs.drain(..) {
            context.define_function(func);
        }

        for (name, code) in compiled.compiled_functions.drain(..) {
            context.set_function_code(&name, code);
        }

        let local_names = compiled.local_names.clone();
        let local_count = compiled.local_count;

        // Run VM and extract locals in a separate scope to release borrow on context
        let (result, vm_locals) = {
            let mut vm = crate::vbscript::vm::Vm::new(context);
            let r = vm.run(compiled);
            let l = std::mem::take(&mut vm.locals);
            (r, l)
        };

        // Write back local variables to context
        for slot in 0..local_count {
            if slot < vm_locals.len() {
                if let Some(name) = local_names.get(slot) {
                    if !name.is_empty() {
                        let val = vm_locals[slot].clone();
                        context.set_variable(name, val);
                    }
                }
            }
        }

        result
    }

    /// Execute multiple ASP blocks with a single VM to preserve variable state.
    /// Converts HTML blocks to Response.Write calls and combines all code.
    pub fn execute_vm_blocks(&self, asp_blocks: &[&AspBlock], context: &mut ExecutionContext) -> Result<(), VBSError> {
        // Convert ASP blocks to a single VBScript code string
        let mut code_parts = Vec::new();
        
        for block in asp_blocks {
            match block {
                AspBlock::Html(html) => {
                    if !html.is_empty() {
                        // Escape the HTML string for VBScript
                        let escaped = html
                            .replace("\"", "\"\"")
                            .replace("\r", "")
                            .replace("\n", "\" & vbCrLf & \"");
                        code_parts.push(format!("Response.Write(\"{}\")", escaped));
                    }
                }
                AspBlock::Code(code, _) => {
                    code_parts.push(code.to_string());
                }
                AspBlock::Directive(_, _) => {
                    // Directives are handled at parse time, ignore here
                }
            }
        }
        
        let combined_code = code_parts.join("\n");
        if combined_code.trim().is_empty() {
            return Ok(());
        }
        
        self.execute_vm(&combined_code, context)
    }

    fn group_tokens_into_lines(&self, tokens: &[Token]) -> Result<Vec<Vec<Token>>, VBSError> {
        let mut lines: Vec<Vec<Token>> = Vec::new();
        let mut current_line: Vec<Token> = Vec::new();
        let mut i = 0;

        while i < tokens.len() {
            let token = &tokens[i];

            match &token.token_type {
                TokenType::NewLine => {
                    if !current_line.is_empty()
                        && !Self::is_line_continuation(&current_line) {
                            lines.push(std::mem::take(&mut current_line));
                        }
                }

                TokenType::WhiteSpace => {
                    if Self::is_continuation_sequence(&tokens[i..]) {
                        i = Self::skip_continuation_sequence(&tokens[i..]);
                    } else if !current_line.is_empty() {
                        current_line.push(token.clone());
                    }
                }

                TokenType::Comment => {
                    current_line.push(token.clone());
                    while i + 1 < tokens.len()
                        && !matches!(tokens[i + 1].token_type, TokenType::NewLine)
                    {
                        i += 1;
                    }
                }

                TokenType::Colon => {
                    current_line.push(token.clone());

                    if !Self::is_in_string_literal(&current_line) {
                        lines.push(std::mem::take(&mut current_line));
                    }
                }

                TokenType::EOF => {
                    break;
                }

                _ => {
                    current_line.push(token.clone());
                }
            }

            i += 1;
        }

        if !current_line.is_empty() {
            lines.push(current_line);
        }

        let lines = lines
            .into_iter()
            .map(Self::trim_whitespace_tokens)
            .filter(|line| !line.is_empty())
            .collect();

        Ok(lines)
    }

    fn is_line_continuation(line: &[Token]) -> bool {
        if let Some(last_token) = line.last() {
            matches!(last_token.token_type, TokenType::WhiteSpace) && last_token.value.contains('_')
        } else {
            false
        }
    }

    fn is_in_string_literal(tokens: &[Token]) -> bool {
        let mut in_string = false;
        for token in tokens {
            if token.token_type == TokenType::StringLiteral { in_string = !in_string }
        }
        in_string
    }

    fn is_continuation_sequence(tokens: &[Token]) -> bool {
        if tokens.is_empty() {
            return false;
        }

        match &tokens[0].token_type {
            TokenType::WhiteSpace => {
                tokens[0].value.contains('_')
                    && tokens
                        .iter()
                        .skip(1)
                        .take_while(|t| matches!(t.token_type, TokenType::WhiteSpace))
                        .any(|t| matches!(t.token_type, TokenType::NewLine))
            }
            _ => false,
        }
    }

    fn skip_continuation_sequence(tokens: &[Token]) -> usize {
        let mut i = 0;

        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        while i < tokens.len() && matches!(tokens[i].token_type, TokenType::WhiteSpace) {
            i += 1;
        }

        if i < tokens.len() && matches!(tokens[i].token_type, TokenType::NewLine) {
            i += 1;
        }

        i
    }

    fn trim_whitespace_tokens(tokens: Vec<Token>) -> Vec<Token> {
        let mut result = tokens;

        let first_non_ws = result
            .iter()
            .position(|t| !matches!(t.token_type, TokenType::WhiteSpace))
            .unwrap_or(result.len());
        result.drain(..first_non_ws);

        while result
            .last()
            .is_some_and(|t| matches!(t.token_type, TokenType::WhiteSpace))
        {
            result.pop();
        }

        result
    }
}

pub fn inject_vbscript_constants(context: &mut ExecutionContext) {
    let constants: Vec<(&str, VBValue)> = vec![
        ("vbCrLf", VBValue::String("\r\n".into())),
        ("vbCr", VBValue::String("\r".into())),
        ("vbLf", VBValue::String("\n".into())),
        ("vbNewLine", VBValue::String("\n".into())),
        ("vbTab", VBValue::String("\t".into())),
        ("vbNullString", VBValue::String("".into())),
        ("vbNull", VBValue::Null),
        ("vbEmpty", VBValue::Empty),
        ("vbObject", VBValue::Empty),
        ("vbTrue", VBValue::Boolean(true)),
        ("vbFalse", VBValue::Boolean(false)),
    ];
    for (name, value) in constants {
        if context.get_variable(name).is_none() {
            context.set_variable(name, value);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_execute_empty_code() {
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(crate::vbscript::store::Store::new());
        let result = VBScriptInterpreter.execute("", &mut ctx);
        assert!(result.is_ok());

        // Whitespace-only
        let result = VBScriptInterpreter.execute("   \n  \t  ", &mut ctx);
        assert!(result.is_ok());
    }

    #[test]
    fn test_execute_simple_assignment() {
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(crate::vbscript::store::Store::new());
        let result = VBScriptInterpreter.execute("x = 42", &mut ctx);
        assert!(result.is_ok());
        let val = ctx.get_variable("x").cloned().unwrap_or(VBValue::Empty);
        match val {
            VBValue::Number(n) => assert!((n - 42.0).abs() < 1e-10),
            _ => panic!("expected Number"),
        }
    }

    #[test]
    fn test_execute_expression() {
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(crate::vbscript::store::Store::new());
        let result = VBScriptInterpreter.execute("x = 2 + 3", &mut ctx);
        assert!(result.is_ok());
        let val = ctx.get_variable("x").cloned().unwrap_or(VBValue::Empty);
        match val {
            VBValue::Number(n) => assert!((n - 5.0).abs() < 1e-10),
            _ => panic!("expected Number"),
        }
    }

    #[test]
    fn test_execute_syntax_error() {
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(crate::vbscript::store::Store::new());
        let result = VBScriptInterpreter.execute("if x then", &mut ctx);
        assert!(result.is_err());
    }

    #[test]
    fn test_vbscript_constants_injected() {
        let mut ctx = ExecutionContext::new();
        inject_vbscript_constants(&mut ctx);

        assert_eq!(ctx.get_variable("vbCrLf"), Some(&VBValue::String("\r\n".into())));
        assert_eq!(ctx.get_variable("vbCr"), Some(&VBValue::String("\r".into())));
        assert_eq!(ctx.get_variable("vbLf"), Some(&VBValue::String("\n".into())));
        assert_eq!(ctx.get_variable("vbNewLine"), Some(&VBValue::String("\n".into())));
        assert_eq!(ctx.get_variable("vbTab"), Some(&VBValue::String("\t".into())));
        assert_eq!(ctx.get_variable("vbNullString"), Some(&VBValue::String("".into())));
        assert_eq!(ctx.get_variable("vbNull"), Some(&VBValue::Null));
        assert_eq!(ctx.get_variable("vbEmpty"), Some(&VBValue::Empty));
        assert_eq!(ctx.get_variable("vbTrue"), Some(&VBValue::Boolean(true)));
        assert_eq!(ctx.get_variable("vbFalse"), Some(&VBValue::Boolean(false)));
    }

    #[test]
    fn test_vbscript_constants_usable_in_code() {
        let mut ctx = ExecutionContext::new();
        ctx.store = Some(crate::vbscript::store::Store::new());
        inject_vbscript_constants(&mut ctx);

        VBScriptInterpreter
            .execute("x = vbCrLf & vbTab & vbNullString", &mut ctx)
            .unwrap();
        let val = ctx.get_variable("x").cloned().unwrap_or(VBValue::Empty);
        assert_eq!(val, VBValue::String("\r\n\t".into()));
    }

    #[test]
    fn test_vbscript_constants_dont_overwrite() {
        let mut ctx = ExecutionContext::new();
        ctx.set_variable("vbCrLf", VBValue::String("custom".into()));
        inject_vbscript_constants(&mut ctx);
        assert_eq!(ctx.get_variable("vbCrLf"), Some(&VBValue::String("custom".into())));
    }
}
