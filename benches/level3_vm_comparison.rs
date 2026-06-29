use criterion::{black_box, criterion_group, criterion_main, Criterion};

use asperger::vbscript::execution_context::ExecutionContext;
use asperger::vbscript::store::Store;
use asperger::vbscript::VBScriptInterpreter;

fn make_ctx() -> ExecutionContext {
    let mut ctx = ExecutionContext::new();
    ctx.store = Some(Store::new());
    ctx
}

// ── Helpers ────────────────────────────────────────────────────────────

fn bench_both(c: &mut Criterion, name: &str, code: &str) {
    let interp = VBScriptInterpreter;

    c.bench_function(&format!("old/{}", name), |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute(black_box(code), &mut ctx).unwrap();
        });
    });

    c.bench_function(&format!("vm/{}", name), |b| {
        b.iter(|| {
            let mut ctx = make_ctx();
            interp.execute_vm(black_box(code), &mut ctx).unwrap();
        });
    });
}

// ── Basic statements ─────────────────────────────────────────────────────

fn bench_literal_assign(c: &mut Criterion) {
    bench_both(c, "literal_assign", "x = 42");
}

fn bench_math_ops(c: &mut Criterion) {
    bench_both(c, "math_ops", "a = 3.14\nb = 2.72\nc = (a + b) * (a - b) / (a * b)");
}

fn bench_string_concat(c: &mut Criterion) {
    bench_both(c, "string_concat_1k", "Dim i, s\ns = \"\"\nFor i = 1 To 1000\ns = s & \"x\"\nNext");
}

fn bench_if_else_2k(c: &mut Criterion) {
    bench_both(c, "if_else_2k", "Dim i, r\nFor i = 1 To 2000\nIf i > 1000 Then\nr = i * 2\nElse\nr = i\nEnd If\nNext");
}

fn bench_for_empty_10k(c: &mut Criterion) {
    bench_both(c, "for_empty_10k", "Dim i\nFor i = 1 To 10000\nNext");
}

fn bench_while_10k(c: &mut Criterion) {
    bench_both(c, "while_10k", "Dim i\ni = 0\nDo While i < 10000\ni = i + 1\nLoop");
}

fn bench_func_call_1k(c: &mut Criterion) {
    bench_both(c, "func_call_1k", "Function add(a, b)\nadd = a + b\nEnd Function\nDim i, r\nFor i = 1 To 1000\nr = add(i, 2)\nNext");
}

fn bench_nested_func_1k(c: &mut Criterion) {
    bench_both(c, "nested_func_1k",
        "Function inner(x)\ninner = x * 2\nEnd Function\n\
         Function outer(x)\nouter = inner(x) + 1\nEnd Function\n\
         Dim i, r\nFor i = 1 To 1000\nr = outer(i)\nNext");
}

fn bench_array_access(c: &mut Criterion) {
    bench_both(c, "array_access",
        "Dim arr\nReDim arr(9)\nDim i\nFor i = 0 To 9\narr(i) = i * 2\nNext\n\
         Dim s\nFor i = 0 To 9\ns = arr(i)\nNext");
}

fn bench_dict_500(c: &mut Criterion) {
    bench_both(c, "dict_500",
        "Dim d, i\nSet d = CreateObject(\"Scripting.Dictionary\")\n\
         For i = 1 To 500\nd.Add CStr(i), i\nNext\nDim v\nv = d(\"100\")");
}

fn bench_select_case_1k(c: &mut Criterion) {
    bench_both(c, "select_case_1k",
        "Dim i, r\nFor i = 1 To 1000\n\
         Select Case i Mod 4\n\
         Case 0: r = \"zero\"\nCase 1: r = \"one\"\n\
         Case 2: r = \"two\"\nCase 3: r = \"three\"\n\
         End Select\nNext");
}

fn bench_recursive_func(c: &mut Criterion) {
    bench_both(c, "recursive_func",
        "Function Factorial(n)\n\
         If n <= 1 Then\nFactorial = 1\nElse\nFactorial = n * Factorial(n - 1)\n\
         End If\nEnd Function\nx = Factorial(10)");
}

criterion_group!(
    name = vm_comparison;
    config = Criterion::default().sample_size(50);
    targets =
        bench_literal_assign,
        bench_math_ops,
        bench_if_else_2k,
        bench_for_empty_10k,
        bench_while_10k,
        bench_func_call_1k,
        bench_nested_func_1k,
        bench_string_concat,
        bench_array_access,
        bench_dict_500,
        bench_select_case_1k,
        bench_recursive_func,
);

criterion_main!(vm_comparison);
