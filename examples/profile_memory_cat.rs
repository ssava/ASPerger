//! Targeted memory profiling — run single workloads, print totals to stdout.
//! Usage: cargo run --example profile_memory_cat --release -- [workload]
//!   workloads: tokenize, assign, math, loops, func, dict, session, cookies, regexp, all

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
    let _ = Tokenizer::tokenize(code);
}

fn profile_assign() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("x = 42\ny = \"hello\"\nz = True", &mut ctx).unwrap();
}

fn profile_math() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("Dim i, s\nFor i = 2 To 100\na = 3.14 * (i + 2.72) / (i - 1.0)\nNext", &mut ctx).unwrap();
}

fn profile_loops() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute("Dim i, s\ns = \"\"\nFor i = 1 To 1000\ns = s & \"x\"\nNext", &mut ctx).unwrap();
}

fn profile_func() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute(
        "Function inner(x)\ninner = x * 2\nEnd Function\n\
         Function outer(x)\nouter = inner(x) + 1\nEnd Function\n\
         Dim i, r\nFor i = 1 To 100\nr = outer(i)\nNext",
        &mut ctx,
    ).unwrap();
}

fn profile_dict() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    ctx.store = Some(Store::new());
    interp.execute(
        "Dim d, i\n\
         Set d = CreateObject(\"Scripting.Dictionary\")\n\
         For i = 1 To 500\nd.Add CStr(i), i\nNext\n\
         Dim v\nv = d(\"100\")",
        &mut ctx,
    ).unwrap();
}

fn profile_session() {
    let store = Store::new();
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Arc::clone(&store));
    ctx.session.id = "prof-session".to_string();
    asperger::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
    let interp = VBScriptInterpreter;
    interp.execute(
        "Dim i\nFor i = 1 To 200\nSession(\"k\" & CStr(i)) = String(100, \"x\")\nNext\n\
         Dim c\nc = Session.Contents.Count",
        &mut ctx,
    ).unwrap();
}

fn profile_cookies() {
    let mut ctx = ExecutionContext::new();
    asperger::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx);
    let interp = VBScriptInterpreter;
    interp.execute(
        "Dim i\nFor i = 1 To 100\n\
         Response.Cookies(\"c\" & CStr(i)) = \"value\" & CStr(i)\n\
         Response.Cookies(\"c\" & CStr(i)).Expires = \"2026-12-31\"\nNext",
        &mut ctx,
    ).unwrap();
}

fn profile_regexp() {
    let interp = VBScriptInterpreter;
    let mut ctx = make_ctx();
    interp.execute(
        "Set re = CreateObject(\"VBScript.RegExp\")\n\
         re.Pattern = \"\\\\d+\"\nre.Global = True\n\
         Dim result\nresult = re.Execute(\"abc 123 def 456 ghi 789\")",
        &mut ctx,
    ).unwrap();
}

fn main() {
    let _dhat = dhat::Dhat::start_heap_profiling();

    let args: Vec<String> = std::env::args().collect();
    let workload = args.get(1).map(|s| s.as_str()).unwrap_or("all");

    // Run one workload at a time so stdout shows the dhat line for just that workload
    match workload {
        "tokenize" => profile_tokenize(),
        "assign" => profile_assign(),
        "math" => profile_math(),
        "loops" => profile_loops(),
        "func" => profile_func(),
        "dict" => profile_dict(),
        "session" => profile_session(),
        "cookies" => profile_cookies(),
        "regexp" => profile_regexp(),
        "all" => {
            profile_tokenize();
            profile_assign();
            profile_math();
            profile_loops();
            profile_func();
            profile_dict();
            profile_session();
            profile_cookies();
            profile_regexp();
        }
        _ => eprintln!("Unknown workload: {workload}"),
    }
}
