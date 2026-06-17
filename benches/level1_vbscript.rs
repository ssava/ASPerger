use criterion::{black_box, criterion_group, criterion_main, Criterion};

use asperger::vbscript::execution_context::ExecutionContext;
use asperger::vbscript::store::Store;
use asperger::vbscript::VBScriptInterpreter;

fn make_ctx() -> ExecutionContext {
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Store::new());
    ctx
}

fn bench_tokenize(c: &mut Criterion) {
    let code = black_box(
        "Dim i, s\nFor i = 1 To 10000\ns = s & \"x\"\nNext",
    );
    c.bench_function("tokenize_short", |b| {
        b.iter(|| {
            let _tokens = asperger::vbscript::Tokenizer::tokenize(code);
        });
    });
}

fn bench_literal_assign(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    c.bench_function("vb_literal_assign", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(black_box("x = 42"), &mut ctx).unwrap();
        });
    });
}

fn bench_math_ops(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("a = 3.14\nb = 2.72\nc = (a + b) * (a - b) / (a * b)");
    c.bench_function("vb_math_ops", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_for_loop_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i, s\nFor i = 1 To 1000\ns = s & \"x\"\nNext");
    c.bench_function("vb_for_loop_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_for_loop_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i\nFor i = 1 To 10000\nNext");
    c.bench_function("vb_for_loop_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_while_loop_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i\ni = 0\nDo While i < 10000\ni = i + 1\nLoop");
    c.bench_function("vb_while_loop_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_function_call(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Function add(a, b)\nadd = a + b\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = add(i, 2)\nNext",
    );
    c.bench_function("vb_function_call_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_array_access(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim arr\nReDim arr(9)\nDim i\nFor i = 0 To 9\narr(i) = i * 2\nNext\nDim s\nFor i = 0 To 9\ns = arr(i)\nNext",
    );
    c.bench_function("vb_array_access", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_string_concat(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i, s\ns = \"\"\nFor i = 1 To 1000\ns = s & \"x\"\nNext");
    c.bench_function("vb_string_concat_1k", |b| {
        b.iter(|| {
            // String concat creates a new string each iteration — measure raw throughput
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_dictionary(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim d, i\nSet d = CreateObject(\"Scripting.Dictionary\")\nFor i = 1 To 500\nd.Add CStr(i), i\nNext\nDim v\nv = d(\"100\")",
    );
    c.bench_function("vb_dictionary_500", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            // inject_asp_intrinsic_objects not needed for CreateObject
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_if_else(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim i, r\nFor i = 1 To 2000\nIf i > 1000 Then\nr = i * 2\nElse\nr = i\nEnd If\nNext",
    );
    c.bench_function("vb_if_else_2k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

criterion_group!(
    name = vbscript;
    config = Criterion::default().sample_size(100);
    targets =
        bench_tokenize,
        bench_literal_assign,
        bench_math_ops,
        bench_for_loop_1k,
        bench_for_loop_10k,
        bench_while_loop_10k,
        bench_function_call,
        bench_array_access,
        bench_string_concat,
        bench_dictionary,
        bench_if_else,
);

criterion_main!(vbscript);
