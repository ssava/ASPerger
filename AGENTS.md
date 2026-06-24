# ASPerger — Agent Guide

## Build & Test
- Build: `cargo build`
- Check: `cargo check`
- Test all: `cargo test --lib`
- Test single: `cargo test --lib <test_name>`
- Lint: `cargo clippy`
- Run server: `cargo run -- asp_files/`
- Run debug adapter: `cargo run --bin asperger-debug -- --help`

## Architecture
- `src/asp/` — HTTP server, ASP block parser, handler chain, error types
- `src/vbscript/` — VBScript interpreter, tokenizer, syntax nodes, built-in functions, COM objects, debugger
- `src/bin/` — Binary entry points (asperger-debug DAP adapter)
- `extension/` — VS Code extension for debugging

## Key Patterns
- `AspBlock` has 3 variants: `Html(String)`, `Code(String, usize)`, `Directive(String, String)`
  - `Code` variant includes the ASP file line number of the first code character (1-indexed)
- Handler chain: `HtmlHandler → CodeHandler` (chain of responsibility)
- `ExecutionContext` owns all per-request state (response buffer, variables, session, etc.)
  - `code_start_line: usize` — physical ASP line where current VBScript block starts (applied as offset in `execute_blocks`)
  - `debugger: Option<Arc<Debugger>>` — shared DAP debugger (must be Sync via `Mutex<Receiver>`)
- `VBScriptObject` trait for ASP intrinsic objects and COM objects
- All ASP intrinsic objects injected as globals before script execution
- Tests live in `src/vbscript/tests.rs` under `#[cfg(test)] mod tests`

## Debugger Line Number Resolution
- `block.line()` returns the logical line within the VBScript code segment (1-indexed)
- `context.code_start_line` is the physical ASP file line of the first code character
- In `execute_blocks`, the check-file-line is computed as:
  - If `code_start_line > 0`: `block.line() + code_start_line - 1`
  - Otherwise: `block.line()` (no offset)
- `execute_user_defined_function` saves/resets `code_start_line` to 0 for function bodies
- Line numbers are computed in `AspParser::parse` by pre-scanning the original content
  - Multi-line blocks (code starts with `\n`): first code line = `<%` line + 1
  - Single-line blocks: first code line = `<%` line

## Debugger Architecture (HTTP Server Mode)
- `asperger_debug.rs` `ConfigurationDone` starts an HTTP server on port 9090
- Each request creates `ExecutionContext` with `Arc::clone(&debugger)` (shared state)
- One request at a time (sequential in a `tokio::runtime::Builder::new_current_thread()`)
- `Debugger` wrapped in `Mutex<Receiver>` to satisfy `Sync` bound for `Arc`

## Breakpoint Flow
1. VS Code sends `SetBreakpoints` → breakpoints stored in `DebuggerState.breakpoints`
2. Browser navigates to `http://127.0.0.1:9090/page.asp`
3. HTTP server reads request, processes ASP blocks through handler chain
4. `CodeHandler::handle` sets `context.code_start_line` from `AspBlock::Code.1`
5. `VBScriptInterpreter::execute` → `execute_blocks` → `debugger.check(file, physical_line, depth)`
6. If line matches a breakpoint → sends `Stopped` event → blocks on `command_rx.recv()`
7. User presses Continue in VS Code → DAP sends `Continue` → `command_rx` unblocks → execution resumes
8. After script finishes → HTTP response sent to browser

## Parser: Dot Precedence
- `Dot` in `precedence()` (expr.rs:92) returns precedence 80 — higher than Concat (35)
- Without this, `"a" & obj.Prop` parses incorrectly as `PropertyAccess(Concat("a", obj), "Prop")` instead of `Concat("a", PropertyAccess(obj, "Prop"))`
- The `prec` variable in `parse_binary` must also handle `TokenType::Dot` specifically (else branch at line 391) to return precedence 80 instead of default `_ => 0`

## Runtime: `eval_binary` Type Check
- `eval_binary` at `expr.rs:769` has a blanket type check rejecting Object/Array for all operators
- `Is`, `Eq`, `Ne`, `Concat` are excluded (can handle objects via `values_equal` or `to_string_val`)
- Without this exclusion, `entry Is Nothing` (where `entry` is a Dictionary) throws `Type mismatch` because `Is` never gets to run its reference comparison

## Common Pitfalls
- `entries.Keys()` (with parens) → method call dispatch, errors "Method 'Keys' not found on Dictionary". Use `entries.Keys` (no parens) for property access
- `obj.Property(args)` in expression context works as `evaluate` tries property+indexed_get first (like `Request.QueryString("id")` or `Application.Contents("key")`)
- `obj(args)` in expression context is parsed as FunctionCall; `evaluate` handles single-arg on Object as indexed_get
- `obj.Property = value` (without parens on obj) is only parsed for `obj.Property = val` — the parser has no general `obj.Property(args) = value` rule. Use `Application("key") = value` instead (matches arr(idx)=expr rule, which dispatches to Object.indexed_set)
- Dictionary.Add uses `HashMap::insert` (overwrites existing keys silently). VBScript's real Dictionary.Add errors on duplicate key
- `Application.Contents("key")` reads return a deep clone of the stored value (Object.clone_box)

## Important Constraints
- `dap` crate v0.4.1-alpha1: response structs in `dap::responses`, event bodies in `dap::events`, `StoppedEventReason` enum in `types`
- `StdoutLock` is `!Send` — use `BufWriter<std::io::Stdout>` for Server writer
- `Command` variants wrap typed args: use `.additional_data` for flattened fields like "program"
- New syntax nodes go in `src/vbscript/syntax/` and re-exported from `mod.rs`
- `AspServer::process_request` now takes `debugger: Option<Arc<Debugger>>` — pass `None` for non-debug mode
- `AspServer::write_response` is now `pub async` (used from debug adapter)
- `contributes.breakpoints` is required in `package.json` separate from `debuggers[].languages`
- VS Code 1.108+ validates DAP response `command` field — wrong body types cause breakpoint disabling
