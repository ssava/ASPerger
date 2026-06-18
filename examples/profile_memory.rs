//! Memory usage profiler using dhat.
//! Runs key ASP/VBScript workloads under heap profiling.
//! Output: dhat-heap.json (viewable at https://nnethercote.github.io/dhat-view/)
//!
//! Usage: cargo run --bin profile_memory --release

use std::sync::Arc;

use asperger::vbscript::execution_context::ExecutionContext;
use asperger::vbscript::store::Store;
use asperger::vbscript::tokenizer::Tokenizer;
use asperger::vbscript::VBScriptInterpreter;

#[global_allocator]
static ALLOC: dhat::DhatAlloc = dhat::DhatAlloc;

fn make_ctx() -> ExecutionContext {
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Store::new());
    ctx
}

fn profile_tokenize() {
    let code = include_str!("../benches/level1_vbscript.rs");
    // Tokenize a large code string to measure token allocation
    let _ = Tokenizer::tokenize(code);
}

fn profile_literal_assign() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("x = 42\ny = \"hello\"\nz = True", &mut ctx).unwrap();
}

fn profile_math_ops() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("a = 3.14\nb = 2.72\nc = (a + b) * (a - b) / (a * b)", &mut ctx).unwrap();
}

fn profile_loops_and_concat() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("Dim i, s\ns = \"\"\nFor i = 1 To 1000\ns = s & \"x\"\nNext", &mut ctx).unwrap();
}

fn profile_function_calls() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute(
        "Function inner(x)\ninner = x * 2\nEnd Function\n\
         Function outer(x)\nouter = inner(x) + 1\nEnd Function\n\
         Dim i, r\nFor i = 1 To 100\nr = outer(i)\nNext",
        &mut ctx,
    ).unwrap();
}

fn profile_dictionary() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    ctx.store = Some(Store::new());
    interp.execute(
        "Dim d, i\n\
         Set d = CreateObject(\"Scripting.Dictionary\")\n\
         For i = 1 To 500\n\
         d.Add CStr(i), i\n\
         Next\n\
         Dim v\nv = d(\"100\")",
        &mut ctx,
    ).unwrap();
}

fn profile_session_vars() {
    let store = Store::new();
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Arc::clone(&store));
    ctx.session.id = "prof-session".to_string();
    asperger::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
    let interp = VBScriptInterpreter;
    // Set 200 session variables with large values
    interp.execute(
        "Dim i\nFor i = 1 To 200\nSession(\"k\" & CStr(i)) = String(100, \"x\")\nNext\n\
         Dim c\nc = Session.Contents.Count\nDim k1, k2\nk1 = Session.Contents.Key(1)\nk2 = Session.Contents.Item(1)",
        &mut ctx,
    ).unwrap();
}

fn profile_response_cookies() {
    let mut ctx = ExecutionContext::new();
    asperger::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
    let interp = VBScriptInterpreter;
    interp.execute(
        "Dim i\n\
         For i = 1 To 100\n\
         Response.Cookies(\"c\" & CStr(i)) = \"value\" & CStr(i)\n\
         Response.Cookies(\"c\" & CStr(i)).Expires = \"2026-12-31\"\n\
         Response.Cookies(\"c\" & CStr(i)).Path = \"/app\"\n\
         Next",
        &mut ctx,
    ).unwrap();
}

fn profile_regexp() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute(
        "Set re = CreateObject(\"VBScript.RegExp\")\n\
         re.Pattern = \"\\\\d+\"\n\
         re.Global = True\n\
         Dim result, m\n\
         result = re.Execute(\"abc 123 def 456 ghi 789\")",
        &mut ctx,
    ).unwrap();
}

fn main() {
    let _dhat = dhat::Dhat::start_heap_profiling();

    // Warm-up (parse-heavy)
    profile_tokenize();

    profile_literal_assign();
    profile_math_ops();
    profile_loops_and_concat();
    profile_function_calls();
    profile_dictionary();
    profile_session_vars();
    profile_response_cookies();
    profile_regexp();

    // Flush and write dhat-heap.json when _dhat drops
}
