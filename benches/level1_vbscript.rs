use criterion::{black_box, criterion_group, criterion_main, Criterion};

use asperger::vbscript::execution_context::ExecutionContext;
use asperger::vbscript::store::Store;
use asperger::vbscript::VBScriptInterpreter;

fn make_ctx() -> ExecutionContext {
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Store::new());
    ctx
}

// ── Tokenizer ────────────────────────────────────────────────────────────

fn bench_tokenize(c: &mut Criterion) {
    let code = black_box("Dim i, s\nFor i = 1 To 10000\ns = s & \"x\"\nNext");
    c.bench_function("tokenize/short", |b| {
        b.iter(|| {
            let _tokens = asperger::vbscript::Tokenizer::tokenize(code);
        });
    });
}

// ── Basic statements ─────────────────────────────────────────────────────

fn bench_literal_assign(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    c.bench_function("stmt/literal_assign", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(black_box("x = 42"), &mut ctx).unwrap();
        });
    });
}

fn bench_math_ops(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("a = 3.14\nb = 2.72\nc = (a + b) * (a - b) / (a * b)");
    c.bench_function("stmt/math_ops", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Loops ────────────────────────────────────────────────────────────────

fn bench_for_empty_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i\nFor i = 1 To 1000\nNext");
    c.bench_function("loop/for_empty_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_for_loop_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i, s\nFor i = 1 To 1000\ns = s & \"x\"\nNext");
    c.bench_function("loop/for_concat_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_for_loop_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i\nFor i = 1 To 10000\nNext");
    c.bench_function("loop/for_empty_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_while_simple_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i\ni = 0\nDo While i < 10000\ni = i + 1\nLoop");
    c.bench_function("loop/while_simple_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_while_empty_condition_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    // Condition is a simple boolean literal — no variable lookup
    let code = black_box("Dim i\ni = 0\nDo While True\ni = i + 1\nIf i >= 10000 Then Exit Do\nLoop");
    c.bench_function("loop/while_true_exit_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_while_string_condition_10k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    // Condition with string comparison (requires type dispatch + to_string conversion)
    let code = black_box("Dim i, s\ni = 0\ns = \"\"\nDo While Len(s) < 10000\ni = i + 1\ns = s & \"x\"\nLoop");
    c.bench_function("loop/while_string_cond_10k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Function calls ───────────────────────────────────────────────────────

fn bench_func_call_0args_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Function f()\nf = 42\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = f()\nNext",
    );
    c.bench_function("func/call_0args_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_func_call_2args_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Function add(a, b)\nadd = a + b\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = add(i, 2)\nNext",
    );
    c.bench_function("func/call_2args_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_func_call_5args_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Function sum(a, b, c, d, e)\nsum = a + b + c + d + e\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = sum(i, 2, 3, 4, 5)\nNext",
    );
    c.bench_function("func/call_5args_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_func_nested_1k(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Function inner(x)\ninner = x * 2\nEnd Function\nFunction outer(x)\nouter = inner(x) + 1\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = outer(i)\nNext",
    );
    c.bench_function("func/nested_call_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_func_call_no_cache(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    // Define and call in same script: each execution redefines the function (parse once cached)
    let code = black_box(
        "Function f(x)\nf = x + 1\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = f(i)\nNext",
    );
    c.bench_function("func/call_redefine_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Data structures ──────────────────────────────────────────────────────

fn bench_array_access(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim arr\nReDim arr(9)\nDim i\nFor i = 0 To 9\narr(i) = i * 2\nNext\nDim s\nFor i = 0 To 9\ns = arr(i)\nNext",
    );
    c.bench_function("data/array_access", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_string_concat(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box("Dim i, s\ns = \"\"\nFor i = 1 To 1000\ns = s & \"x\"\nNext");
    c.bench_function("data/string_concat_1k", |b| {
        b.iter(|| {
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
    c.bench_function("data/dictionary_500", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Control flow ─────────────────────────────────────────────────────────

fn bench_if_else(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim i, r\nFor i = 1 To 2000\nIf i > 1000 Then\nr = i * 2\nElse\nr = i\nEnd If\nNext",
    );
    c.bench_function("ctrl/if_else_2k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_select_case(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim i, r\nFor i = 1 To 1000\nSelect Case i Mod 4\nCase 0: r = \"zero\"\nCase 1: r = \"one\"\nCase 2: r = \"two\"\nCase 3: r = \"three\"\nEnd Select\nNext",
    );
    c.bench_function("ctrl/select_case_1k", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Date functions (new: string date support) ──────────────────────────

fn bench_date_functions_ole(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim d, y, m, dy\nd = DateSerial(2026, 6, 19)\ny = Year(d)\nm = Month(d)\ndy = Day(d)",
    );
    c.bench_function("date/ole_serial", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_date_functions_string(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim d, y, m, dy\nd = CDate(\"06/19/2026\")\ny = Year(d)\nm = Month(d)\ndy = Day(d)",
    );
    c.bench_function("date/from_cdate_string", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_date_functions_raw_string(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Dim y, m, dy\ny = Year(\"2026-06-19\")\nm = Month(\"2026-06-19\")\ndy = Day(\"2026-06-19\")",
    );
    c.bench_function("date/raw_string", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

// ── Class methods (new: Sub/Function dispatch) ─────────────────────────

fn bench_class_method_no_args(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Class Counter\nPublic count\nPublic Sub Inc\ncount = count + 1\nEnd Sub\nEnd Class\n\
         Dim c, i\nSet c = New Counter\nc.count = 0\nFor i = 1 To 100\nc.Inc\nNext",
    );
    c.bench_function("class/method_no_args_100", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_class_function_return(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Class Calc\nPublic Function Add(a, b)\nAdd = a + b\nEnd Function\nEnd Class\n\
         Dim c, i, r\nSet c = New Calc\nFor i = 1 To 100\nr = c.Add(i, 2)\nNext",
    );
    c.bench_function("class/function_return_100", |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(code, &mut ctx).unwrap();
        });
    });
}

fn bench_class_method_mutates_instance(c: &mut Criterion) {
    let interp = VBScriptInterpreter;
    let code = black_box(
        "Class Accum\nPublic total\nPublic Sub Add(n)\ntotal = total + n\nEnd Sub\nEnd Class\n\
         Dim a, i\nSet a = New Accum\na.total = 0\nFor i = 1 To 100\na.Add(i)\nNext",
    );
    c.bench_function("class/instance_mutation_100", |b| {
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
        bench_for_empty_1k,
        bench_for_loop_1k,
        bench_for_loop_10k,
        bench_while_simple_10k,
        bench_while_empty_condition_10k,
        bench_while_string_condition_10k,
        bench_func_call_0args_1k,
        bench_func_call_2args_1k,
        bench_func_call_5args_1k,
        bench_func_nested_1k,
        bench_func_call_no_cache,
        bench_array_access,
        bench_string_concat,
        bench_dictionary,
        bench_if_else,
        bench_select_case,
        bench_date_functions_ole,
        bench_date_functions_string,
        bench_date_functions_raw_string,
        bench_class_method_no_args,
        bench_class_function_return,
        bench_class_method_mutates_instance,
);

criterion_main!(vbscript);
