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

## VM Benchmarks (level3_vm_comparison)
- Run: `cargo bench --bench level3_vm_comparison` (12 benchmarks, 50 samples each)
- VM is **27-69% faster** on loop-heavy workloads vs old interpreter
- Trivial single-statement benchmarks are slower (~30-40%) due to compilation overhead
- Function calls bridge to old interpreter (slightly slower for now)
- Key wins: `for_empty_10k` (+69%), `if_else_2k` (+43%), `while_10k` (+27%), `select_case_1k` (+21%), `dict_500` (+21%)

## VM: Array Indexing
- Arrays are 0-based in VBScript (default `Option Base 0`)
- `ReDim arr(9)` allocates 10 elements (indices 0-9): `total_size *= bound + 1`
- `IndexGet`, `IndexStoreLocal`, `IndexStoreGlobal` use `idx = to_arg_f64(&key) as usize` (0-based, no `- 1`)
- `Arc::make_mut` for copy-on-write array mutation (replaces `Arc::get_mut` which fails on shared refs)

## VM: Object Method Call / Property Set Mutation
- `CallMethod` operates on a **clone** of the object — mutations to `self` inside `call_method` are lost
- `CallMethodLocal(LocalSlot, ConstantIdx, u8)` and `CallMethodGlobal(ConstantIdx, ConstantIdx, u8)` swap the object out of the variable, call the method, and swap it back
- `SetProp(ConstantIdx)` operates on a **clone** of the object — mutations to `self` inside `set_property` are lost
- `SetPropLocal(LocalSlot, ConstantIdx)` and `SetPropGlobal(ConstantIdx, ConstantIdx)` swap the object out of the variable, set the property, and swap it back
- Used by `PropertySet::compile` for `obj.Property = value` statements (e.g. `c.count = 5`)
- `CallMethod` (clone) works for read-only methods and ASP intrinsics (which mutate via `context`, not `self`)
- `CallMethodGlobal`/`CallMethodLocal` should NOT wrap errors from `call_method` with `map_err` — this corrupts error codes from `Err.Raise`

## VM: FunctionCall Array/Object Indexing
- `Expr::FunctionCall { name, args }` in compiler checks if `name` is a local variable → emits `LoadLocal(slot)` + `IndexGet` (array read)
- `Instruction::Call` handler checks array access → object indexed access → user function → builtin (matching old interpreter's evaluate order)
- Object indexed access swaps the variable out, calls `indexed_get`, swaps it back

## Critical VM Files
- `src/vbscript/vm.rs` — ~1150 lines, `Vm` struct with `execute_loop`, ForState, ForEachState
- `src/vbscript/compiler.rs` — ~680 lines, `Compiler` struct with `compile`, `compile_expr`, `compile_block`
- `src/vbscript/instruction.rs` — ~55 `Instruction` variants + Display impl
- `src/vbscript/syntax/*.rs` — each syntax node has `compile(&self, compiler)` implementing `VBSyntax`
- `benches/level3_vm_comparison.rs` — Old interpreter vs VM benchmarks

## Important Constraints
- `dap` crate v0.4.1-alpha1: response structs in `dap::responses`, event bodies in `dap::events`, `StoppedEventReason` enum in `types`
- `StdoutLock` is `!Send` — use `BufWriter<std::io::Stdout>` for Server writer
- `Command` variants wrap typed args: use `.additional_data` for flattened fields like "program"
- New syntax nodes go in `src/vbscript/syntax/` and re-exported from `mod.rs`
- `AspServer::process_request` now takes `debugger: Option<Arc<Debugger>>` — pass `None` for non-debug mode
- `AspServer::write_response` is now `pub async` (used from debug adapter)
- `contributes.breakpoints` is required in `package.json` separate from `debuggers[].languages`
- VS Code 1.108+ validates DAP response `command` field — wrong body types cause breakpoint disabling
