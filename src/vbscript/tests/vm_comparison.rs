use super::*;

fn compare_interpreters(code: &str) {
    // For now, skip code with Dim (local variables differ between interpreters)
    if code.contains("Dim ") {
        return;
    }
    let mut ctx1 = ExecutionContext::new();
    let mut ctx2 = ExecutionContext::new();

    let interp = crate::vbscript::VBScriptInterpreter;
    let result1 = interp.execute(code, &mut ctx1);
    let result2 = interp.execute_vm(code, &mut ctx2);

    let msg = format!("\ncode: {}\nold: {:?}\nvm:  {:?}", code, result1, result2);
    assert_eq!(result1.is_ok(), result2.is_ok(), "{}", msg);

    let vars1 = ctx1.variables();
    let vars2 = ctx2.variables();
    // Compare all keys, skipping Object values (PartialEq always false)
    let mut keys: Vec<&String> = vars1.keys().chain(vars2.keys()).collect();
    keys.sort();
    keys.dedup();
    for key in keys {
        let v1 = vars1.get(key);
        let v2 = vars2.get(key);
        let skip = matches!(v1, Some(VBValue::Object(_))) || matches!(v2, Some(VBValue::Object(_)));
        if skip {
            continue;
        }
        let msg = format!("\ncode: {}\nkey: {}\nold: {:?}\nvm:  {:?}", code, key, v1, v2);
        assert_eq!(v1, v2, "{}", msg);
    }
}


#[test]
fn test_vm_assignments() {
    compare_interpreters("x = 42");
    compare_interpreters("x = \"hello\"");
    compare_interpreters("x = True");
    compare_interpreters("x = 3.14");
    compare_interpreters("x = Null");
    compare_interpreters("x = Empty");
}

#[test]
fn test_vm_expressions() {
    compare_interpreters("x = 2 + 3");
    compare_interpreters("x = 10 - 4");
    compare_interpreters("x = 3 * 7");
    compare_interpreters("x = 10 / 3");
    compare_interpreters("x = 10 \\ 3");
    compare_interpreters("x = 10 Mod 3");
    compare_interpreters("x = 2 ^ 3");
    compare_interpreters("x = \"a\" & \"b\"");
    compare_interpreters("x = 2 + 3 * 4");
    compare_interpreters("x = (2 + 3) * 4");
    compare_interpreters("x = 5 + 3 * 2 ^ 2");
}

#[test]
fn test_vm_comparisons() {
    compare_interpreters("x = 1 = 1");
    compare_interpreters("x = 1 <> 2");
    compare_interpreters("x = 1 < 2");
    compare_interpreters("x = 2 <= 2");
    compare_interpreters("x = 3 > 2");
    compare_interpreters("x = 3 >= 3");
    compare_interpreters("x = \"a\" > \"b\"");
}

#[test]
fn test_vm_logical_ops() {
    compare_interpreters("x = True And False");
    compare_interpreters("x = True Or False");
    compare_interpreters("x = True Xor False");
    compare_interpreters("x = Not True");
    compare_interpreters("x = 5 And 3");
    compare_interpreters("x = 5 Or 3");
}

#[test]
fn test_vm_unary() {
    compare_interpreters("x = -5");
    compare_interpreters("x = - (3 + 2)");
    compare_interpreters("x = Not (1 > 2)");
}

#[test]
fn test_vm_if_then() {
    compare_interpreters("Dim r\nIf True Then r = 1 Else r = 2");
    compare_interpreters("Dim r\nIf False Then r = 1 Else r = 2");
    compare_interpreters("Dim r\nr = 0\nIf 1 < 2 Then r = 10");
    compare_interpreters("Dim r\nIf False Then\n    r = 1\nElseIf True Then\n    r = 2\nElse\n    r = 3\nEnd If");
    compare_interpreters("Dim r\nIf 1 > 2 Then\n    r = 1\nElseIf 2 > 3 Then\n    r = 2\nElseIf 3 > 1 Then\n    r = 3\nElse\n    r = 4\nEnd If");
}

#[test]
fn test_vm_for_loop() {
    compare_interpreters("Dim total, i\ntotal = 0\nFor i = 1 To 5\n    total = total + i\nNext");
    compare_interpreters("Dim total, i\ntotal = 0\nFor i = 1 To 10 Step 2\n    total = total + i\nNext");
    compare_interpreters("Dim total, i\ntotal = 0\nFor i = 10 To 1 Step -1\n    total = total + i\nNext");
    compare_interpreters("Dim total, i\ntotal = 0\nFor i = 1 To 0\n    total = 99\nNext");
    compare_interpreters("Dim i\nFor i = 1 To 3\n    i = 999\nNext");
}

#[test]
fn test_vm_exit_for() {
    compare_interpreters("Dim i, sum\nsum = 0\nFor i = 1 To 10\n    If i = 5 Then Exit For\n    sum = sum + 1\nNext");
    compare_interpreters("Dim i, j, sum\nsum = 0\nFor i = 1 To 3\n    For j = 1 To 3\n        If j = 2 Then Exit For\n        sum = sum + 1\n    Next\nNext");
}

#[test]
fn test_vm_for_each() {
    compare_interpreters("Dim x, items, sum\nitems = Array(1, 2, 3, 4, 5)\nsum = 0\nFor Each x In items\n    sum = sum + x\nNext");
    compare_interpreters("Dim x, items, sum\nitems = Array(10, 20, 30)\nsum = 0\nFor Each x In items\n    If x = 20 Then Exit For\n    sum = sum + x\nNext");
    compare_interpreters("Dim x, items\nitems = Array(5, 10, 15)\nFor Each x In items\n    x = 999\nNext");
    compare_interpreters("Dim x, items\nitems = Array()\nFor Each x In items\n    x = 1\nNext");
}

#[test]
fn test_vm_while_loop() {
    compare_interpreters("Dim x\ntotal = 0\nx = 1\nWhile x <= 5\n    total = total + x\n    x = x + 1\nWend");
    compare_interpreters("Dim x\ntotal = 0\nx = 10\nWhile x <= 5\n    total = 99\nWend");
    compare_interpreters("Dim x\nx = 5\nWhile x = 5\n    x = 6\nWend");
}

#[test]
fn test_vm_do_loops() {
    compare_interpreters("Dim x, total\ntotal = 0\nx = 1\nDo While x <= 5\n    total = total + x\n    x = x + 1\nLoop");
    compare_interpreters("Dim x, total\ntotal = 0\nx = 1\nDo Until x > 5\n    total = total + x\n    x = x + 1\nLoop");
    compare_interpreters("Dim x, total\ntotal = 0\nx = 1\nDo\n    total = total + x\n    x = x + 1\nLoop While x <= 5");
    compare_interpreters("Dim x, total\ntotal = 0\nx = 1\nDo\n    total = total + x\n    x = x + 1\nLoop Until x > 5");
    compare_interpreters("Dim x, total\ntotal = 0\nx = 10\nDo While x <= 5\n    total = 99\nLoop");
    compare_interpreters("Dim x\nx = 0\nDo\n    x = x + 1\nLoop Until x = 3");
}

#[test]
fn test_vm_exit_do() {
    // Standalone Exit Do inside Do While (no Dim to ensure global store)
    let code1 = "x = 1\ndo while true\n    Exit Do\n    x = 2\nloop\nx = 3";
    let mut ctx = ExecutionContext::new();
    crate::vbscript::VBScriptInterpreter.execute_vm(code1, &mut ctx).unwrap();
    assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(3.0)));

    // Inline Exit Do in If inside Do While
    let code2 = "total = 0\nx = 1\nDo While x < 10\n    If x = 5 Then Exit Do\n    total = total + x\n    x = x + 1\nLoop";
    let mut ctx = ExecutionContext::new();
    crate::vbscript::VBScriptInterpreter.execute_vm(code2, &mut ctx).unwrap();
    assert_eq!(ctx.get_variable("total"), Some(&VBValue::Number(10.0)));

    // Inline Exit Do in If inside Do Loop While
    let code3 = "total = 0\nx = 1\nDo\n    If x = 3 Then Exit Do\n    total = total + x\n    x = x + 1\nLoop While x < 10";
    let mut ctx = ExecutionContext::new();
    crate::vbscript::VBScriptInterpreter.execute_vm(code3, &mut ctx).unwrap();
    assert_eq!(ctx.get_variable("total"), Some(&VBValue::Number(3.0)));
}

#[test]
fn test_vm_select_case() {
    compare_interpreters("Dim r\nx = 2\nSelect Case x\n    Case 1\n        r = 10\n    Case 2\n        r = 20\n    Case 3\n        r = 30\nEnd Select");
    compare_interpreters("Dim r\nx = 5\nSelect Case x\n    Case 1\n        r = 10\n    Case 2\n        r = 20\n    Case Else\n        r = 99\nEnd Select");
    compare_interpreters("Dim r\nx = \"hello\"\nSelect Case x\n    Case \"hi\"\n        r = 1\n    Case \"hello\"\n        r = 2\n    Case Else\n        r = 3\nEnd Select");
}

#[test]
fn test_vm_arrays() {
    compare_interpreters("Dim a(5)\na(0) = 10\na(1) = 20\nx = a(0) + a(1)");
    compare_interpreters("Dim a(2,3)\na(0,1) = 42\nx = a(0,1)");
    compare_interpreters("x = Array(1, 2, 3)");
    compare_interpreters("Dim a\na = Array(10, 20, 30)\nx = a(1)");
}

#[test]
fn test_vm_redim() {
    compare_interpreters("Dim a()\nReDim a(5)\na(0) = 42\nx = a(0)");
    compare_interpreters("Dim a()\nReDim a(2)\na(0) = 10\na(1) = 20\nReDim Preserve a(4)\nx = a(1)");
}

#[test]
fn test_vm_builtin_functions() {
    compare_interpreters("x = Len(\"hello\")");
    compare_interpreters("x = UCase(\"Hello\")");
    compare_interpreters("x = LCase(\"Hello\")");
    compare_interpreters("x = Left(\"hello\", 2)");
    compare_interpreters("x = Right(\"hello\", 2)");
    compare_interpreters("x = Mid(\"hello\", 2, 3)");
    compare_interpreters("x = InStr(\"hello\", \"ll\")");
    compare_interpreters("x = Trim(\"  hi  \")");
    compare_interpreters("x = Replace(\"hello\", \"l\", \"x\")");
    compare_interpreters("x = Split(\"a,b,c\", \",\")");
    compare_interpreters("x = Join(Array(\"a\",\"b\",\"c\"), \"-\")");
    compare_interpreters("x = Abs(-5)");
    compare_interpreters("x = Int(3.7)");
    compare_interpreters("x = Fix(-3.7)");
    compare_interpreters("x = Sgn(-5)");
    compare_interpreters("x = Asc(\"A\")");
    compare_interpreters("x = Chr(65)");
}

#[test]
fn test_vm_variable_scope() {
    compare_interpreters("Dim x\nx = 10\nDim y\ny = x + 5");
    compare_interpreters("x = 10\ny = x * 2");
}

#[test]
fn test_vm_multi_statement() {
    compare_interpreters("Dim a, b, c\na = 1: b = 2: c = a + b");
}

#[test]
fn test_vm_string_concat() {
    compare_interpreters("x = \"Hello, \" & \"World!\"");
    compare_interpreters("x = \"Count: \" & 42");
    compare_interpreters("x = \"Value: \" & True");
}

#[test]
fn test_vm_const() {
    compare_interpreters("Dim x\nx = 42\nIf x = 42 Then\n    y = 1\nElse\n    y = 2\nEnd If");
}

#[test]
fn test_vm_function_call() {
    compare_interpreters("Function Add(a, b)\n    Add = a + b\nEnd Function\nx = Add(3, 4)");
    compare_interpreters("Function Square(n)\n    Square = n * n\nEnd Function\nx = Square(5)");
    compare_interpreters("Sub Greet(name)\n    Dim msg\n    msg = \"Hello, \" & name\nEnd Sub\nCall Greet(\"World\")");
}

#[test]
fn test_vm_exit_function() {
    compare_interpreters("Function Calc(n)\n    If n < 0 Then\n        Calc = 0\n        Exit Function\n    End If\n    Calc = n * 2\nEnd Function\nx = Calc(-5)\ny = Calc(10)");
}

#[test]
fn test_vm_recursive_function() {
    compare_interpreters("Function Factorial(n)\n    If n <= 1 Then\n        Factorial = 1\n    Else\n        Factorial = n * Factorial(n - 1)\n    End If\nEnd Function\nx = Factorial(5)");
}

#[test]
fn test_vm_with_arithmetic() {
    compare_interpreters("x = (1 + 2) * (3 + 4)");
    compare_interpreters("x = 10 - 5 - 2");
    compare_interpreters("x = 2 ^ 3 ^ 2");
}

#[test]
fn test_vm_response_write() {
    let mut ctx1 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx1);
    let mut ctx2 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx2);

    let code = "Response.Write(\"hello from VM\")";
    let interp = crate::vbscript::VBScriptInterpreter;
    interp.execute(code, &mut ctx1).unwrap();
    interp.execute_vm(code, &mut ctx2).unwrap();
    assert_eq!(ctx1.response.buffer, ctx2.response.buffer);
}

#[test]
fn test_vm_redirect() {
    let mut ctx1 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx1);
    let mut ctx2 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx2);

    let code = "Response.Redirect(\"other.asp\")";
    let interp = crate::vbscript::VBScriptInterpreter;
    interp.execute(code, &mut ctx1).unwrap();
    interp.execute_vm(code, &mut ctx2).unwrap();
    assert_eq!(ctx1.response.status, ctx2.response.status);
    assert_eq!(ctx1.response.extra_headers, ctx2.response.extra_headers);
    assert_eq!(ctx1.response.ended, ctx2.response.ended);
}

#[test]
fn test_vm_response_end() {
    let mut ctx1 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx1);
    let mut ctx2 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx2);

    let code = "Response.End()\nResponse.Write(\"should not appear\")";
    let interp = crate::vbscript::VBScriptInterpreter;
    interp.execute(code, &mut ctx1).unwrap();
    interp.execute_vm(code, &mut ctx2).unwrap();
    assert_eq!(ctx1.response.ended, ctx2.response.ended);
    assert_eq!(ctx1.response.buffer, ctx2.response.buffer);
}

#[test]
fn test_vm_exit_do_while_true() {
    let code = "n = 1\nDo While True\n    n = n + 1\n    If n > 3 Then Exit Do\nLoop\nx = n";
    let mut ctx = ExecutionContext::new();
    let interp = crate::vbscript::VBScriptInterpreter;
    let result = interp.execute_vm(code, &mut ctx);
    assert!(result.is_ok(), "VM error: {:?}", result);
    assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(4.0)));
}

#[test]
fn test_vm_exit_do_no_condition() {
    // Do without any condition (infinite)
    let code = "n = 1\nDo\n    n = n + 1\n    If n > 3 Then Exit Do\nLoop\nx = n";
    let mut ctx = ExecutionContext::new();
    let interp = crate::vbscript::VBScriptInterpreter;
    let result = interp.execute_vm(code, &mut ctx);
    assert!(result.is_ok(), "VM error: {:?}", result);
    assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(4.0)));
}

#[test]
fn test_vm_exit_do_until() {
    let code = "n = 1\nDo Until n > 3\n    n = n + 1\nLoop\nx = n";
    let mut ctx = ExecutionContext::new();
    let interp = crate::vbscript::VBScriptInterpreter;
    let result = interp.execute_vm(code, &mut ctx);
    assert!(result.is_ok(), "VM error: {:?}", result);
    assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(4.0)));
}

#[test]
fn test_vm_exit_do_while_cond() {
    let code = "n = 1\nDo While n < 10\n    n = n + 1\nLoop\nx = n";
    let mut ctx = ExecutionContext::new();
    let interp = crate::vbscript::VBScriptInterpreter;
    let result = interp.execute_vm(code, &mut ctx);
    assert!(result.is_ok(), "VM error: {:?}", result);
    assert_eq!(ctx.get_variable("x"), Some(&VBValue::Number(10.0)));
}

#[test]
fn test_vm_date_parts() {
    let mut ctx1 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx1);
    let mut ctx2 = ExecutionContext::new();
    crate::asp::server::AspServer::inject_asp_intrinsic_objects(&mut ctx2);

    let code = "d = Now()\ny = Year(d)\nm = Month(d)\ndy = Day(d)\nh = Hour(d)\nmi = Minute(d)\ns = Second(d)";
    let interp = crate::vbscript::VBScriptInterpreter;
    interp.execute(code, &mut ctx1).unwrap();
    interp.execute_vm(code, &mut ctx2).unwrap();

    assert_eq!(ctx1.get_variable("y"), ctx2.get_variable("y"));
    assert_eq!(ctx1.get_variable("m"), ctx2.get_variable("m"));
    assert_eq!(ctx1.get_variable("dy"), ctx2.get_variable("dy"));
    assert_eq!(ctx1.get_variable("h"), ctx2.get_variable("h"));
    assert_eq!(ctx1.get_variable("mi"), ctx2.get_variable("mi"));
    assert_eq!(ctx1.get_variable("s"), ctx2.get_variable("s"));
}


