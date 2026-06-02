# ASPerger

A lightweight **ASP Classic / VBScript** server written in Rust. Parses and executes `.asp` files, serves them over HTTP, and supports step-through debugging in VS Code.

## Features

- **HTTP server** — serves `.asp` files with `<% %>` code blocks + static files
- **Full VBScript interpreter** — custom tokenizer, Pratt expression parser, block evaluator
- **Preprocessor** — `<!-- #include file="..." -->` / `<!-- #include virtual="..." -->` with recursive expansion and cycle detection; `<%@ LANGUAGE %`, `<%@ ENABLESESSIONSTATE %>`, `<%@ CODEPAGE %>`, `<%@ LCID %>`, `<%@ TRANSACTION %>` directives
- **Control flow** — `If/Then/ElseIf/Else/End If`, `For/Next`, `For Each/Next`, `While/Wend`, `Do/Loop` (pre/post-test, While/Until), `Select Case`
- **Functions & Subs** — `Function`/`End Function`, `Sub`/`End Sub`, `Call`, `Exit Function`, `Exit Sub`, `Exit For`, `Exit Do`
- **Classes** — `Class`/`End Class` with `Public`/`Private` members, `Property Get`/`Property Let`, `With`/`End With`
- **Operators** — arithmetic (`+`, `-`, `*`, `/`, `\`, `^`, `Mod`), comparison (`=`, `<>`, `<`, `>`, `<=`, `>=`, `Is`), logical (`And`, `Or`, `Not`, `Xor`, `Eqv`, `Imp`), string concat (`&`, `+`)
- **Error handling** — `On Error Resume Next` / `On Error Goto 0`, `Err.Raise`, `Err.Number`, `Err.Description`, `Err.Source`
- **Arrays** — `Array()`, `ReDim`, `ReDim Preserve`, `Split`, `Join`, `Filter`, `LBound`, `UBound`
- **Date/Time** — `Now`, `Date`, `Time`, `DateAdd`, `DateDiff`, `DateSerial`, `TimeSerial`, `DateValue`, `TimeValue`, `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`, `Weekday`, `WeekdayName`, `MonthName`, `FormatDateTime`, `Timer`
- **Math** — `Abs`, `Sgn`, `Sqr`, `Int`, `Fix`, `Round`, `Rnd`, `Hex`, `Oct`
- **Type checks** — `IsArray`, `IsDate`, `IsEmpty`, `IsNull`, `IsNumeric`, `IsObject`, `VarType`, `TypeName`
- **Type conversions** — `CInt`, `CLng`, `CBool`, `CByte`, `CDbl`, `CDate`, `CStr`, `Hex`, `Oct`
- **String functions** — `Len`, `Mid`, `Left`, `Right`, `Trim`/`LTrim`/`RTrim`, `UCase`/`LCase`, `InStr`/`InStrRev`, `Replace`, `Split`/`Join`, `Asc`/`Chr`, `Space`/`String`, `StrReverse`, `StrComp`, `Filter`, `FormatCurrency`/`FormatNumber`/`FormatPercent`, `LSet`/`RSet`

### ASP Intrinsic Objects

| Object | Status | Key members |
|--------|--------|-------------|
| `Request` | ✅ | `Form`, `QueryString`, `Cookies`, `ServerVariables`, `TotalBytes` — all with `.Count` |
| `Response` | ✅ | `.Write()`, `.End()`, `.Buffer`, `.ContentType`, `.Status`, `.Expires`, `.Cookies` |
| `Session` | ✅ | `.SessionID`, `.Timeout`, `.Abandon()`, `.Contents.Count`, indexed `Session("key")` — disabled when `<%@ ENABLESESSIONSTATE=False %>` |
| `Server` | ✅ | `.HTMLEncode()`, `.URLEncode()`, `.URLPathEncode()`, `.MapPath()`, `.CreateObject()`, `.ScriptTimeout`, `.ScriptPath`, `.Execute()`, `.Transfer()` |
| `Application` | ✅ | `.Lock()`/`.Unlock()`, `.Contents.Count`, indexed `Application("key")` |

### COM Objects

| Object | Status | Key members |
|--------|--------|-------------|
| `Scripting.Dictionary` | ✅ | `.Count`, `.Keys`, `.Items`, `.Add()`, `.Remove()`, `.Exists()`, `.RemoveAll()`, indexed access |
| `RegExp` | ✅ | `.Pattern`, `.IgnoreCase`, `.Global`, `.Test()`, `.Execute()`, `.Replace()` |
| `ADODB.Connection` | ✅ | `.ConnectionString`, `.Open()`, `.Close()`, `.Execute()` → Recordset |
| `ADODB.Recordset` | ✅ | `.EOF`, `.MoveNext()`, `.Fields("name").Value` |
| `Scripting.FileSystemObject` | ✅ | `.CreateTextFile()`, `.OpenTextFile()`, `.FileExists()`, `.FolderExists()`, `.GetFile()`, `.GetFolder()`, `.GetAbsolutePathName()`, `.GetSpecialFolder()`, `.CreateFolder()`, `.DeleteFolder()`, `.CopyFolder()`, `.MoveFolder()`, `.DeleteFile()`, `.CopyFile()`, `.MoveFile()` |
| `Scripting.TextStream` | ✅ | `.Read()`, `.ReadLine()`, `.ReadAll()`, `.Write()`, `.WriteLine()`, `.WriteBlankLines()`, `.Skip()`, `.SkipLine()`, `.Close()`, `.AtEndOfStream` |

### Debugging (VS Code)

Full Debug Adapter Protocol (DAP) support — step through VBScript code in VS Code:

- Step over / Step in / Step out / Continue / Pause
- Breakpoints by file + line
- Local variable inspection (name, value, type)
- Call stack with frame names and locations
- Runtime error output to debug console

See [`DEBUG.md`](DEBUG.md) for setup instructions.

## Quick start

```bash
# Build
cargo build --release

# Run
cargo run -- --port 9090 --folder ./asp_files
```

Then open `http://localhost:9090` in your browser.

## CLI options

| Option | Default | Description |
|--------|---------|-------------|
| `--host` | `127.0.0.1` | Bind address |
| `-p`, `--port` | `8080` | Port number |
| `-f`, `--folder` | `./` | Directory containing ASP files |

Example:
```bash
cargo run -- --host 0.0.0.0 --port 3000 --folder ./www
```

## Demo

The repository includes a self-evaluating test suite at `asp_files/index.asp` that exercises **all 29 tests** and displays a **Summary: 29/29 passed** at the bottom. Run the server and open the page to see:

- `Response.Write` (with and without parentheses)
- Variable declaration (`Dim`) and assignment (string, numeric, boolean)
- `If/Then/ElseIf/Else/End If` with `And`, `Or`, `Not`
- `For i = 1 To 5 ... Next` loops (including `Step` and negative `Step`)
- `For Each ... In ... Next` with `Scripting.Dictionary`
- `While ... Wend` and `Do ... Loop` (pre/post-test, While/Until)
- `Select Case` with multiple cases and `Case Else`
- User-defined `Function` / `Sub` / `Call` (including `Exit Function`)
- `Array`, `ReDim`, `ReDim Preserve`, `Split`, `Join`, `UBound`
- `Class` / `End Class` with `Property Get` / `Property Let` / `Property Set`
- `With` / `End With` (property get/set, method call)
- `On Error Resume Next` / `On Error Goto 0`, `Err.Raise`
- `Application.Lock` / `.Unlock`, `Application.Contents.Count`
- `Session.SessionID`, `.Timeout`, `.Contents.Count`
- `Response.Buffer`, `.ContentType`, `.Status`, `.Expires`, `.Cookies`
- `Request.Form`, `.QueryString`, `.ServerVariables` (with `.Count`)
- `Server.HTMLEncode`, `.URLEncode`, `.MapPath`, `.ScriptTimeout`, `.ScriptPath`, `.CreateObject`
- `RegExp` — `Test()`, `Execute()`, `Replace()`, `.IgnoreCase`, `.Global`
- `Scripting.Dictionary` — `.Count`, `.Keys()`, `.Items()`, `.Exists()`, indexed access
- `Scripting.FileSystemObject` — `.FileExists()`, `.FolderExists()`, `.GetAbsolutePathName()`
- Built-in functions: `Len`, `UCase`, `LCase`, `Mid`, `Left`, `Right`, `Trim`, `InStr`, `Replace`, `Split`, `Join`, `Now`, `Date`, `Time`, `DateAdd`, `DateDiff`, `Abs`, `Sqr`, `Sgn`, `Int`, `Fix`, `Round`, `Rnd`, `Hex`, `Oct`, `IsArray`, `IsDate`, `IsEmpty`, `IsNull`, `IsNumeric`, `IsObject`, `VarType`, `TypeName`, `CInt`, `CLng`, `CBool`, `CByte`, `CDbl`, `CDate`, `Asc`, `Chr`, `Space`, `String`, `StrReverse`, `StrComp`, `Filter`, `FormatCurrency`, `FormatNumber`, `FormatPercent`, `LSet`, `RSet`
- String concatenation (`&` and `+`), `Mod`, integer division (`\`), `Eqv`, `Imp`
- `Empty`, `Null`, `Nothing`, `IsEmpty`/`IsNull`
- Date literals (`#2024-01-15#`), HEX literals (`&HFF`), octal literals (`&77`)
- Line continuation (`_`), comment (`'` / `REM`), statement separator (`:`)

## VBScript support

| Category | Status |
|----------|--------|
| Variables (`Dim`, `Set`, assignment) | ✅ |
| Control flow (`If`, `For`, `While`, `Do`, `Select Case`) | ✅ |
| `Function` / `Sub` / `Call` (including `Exit`) | ✅ |
| `Class` / `Property Get/Let/Set` / `With` | ✅ |
| Arrays (`Array()`, `ReDim`, `ReDim Preserve`, index) | ✅ |
| Built-in functions (50+ string, math, date, type, array) | ✅ |
| `Err` object / `Err.Raise` / `On Error Resume Next \| Goto 0` | ✅ |
| `Scripting.Dictionary` | ✅ |
| `RegExp` | ✅ |
| `ADODB.Connection` + `Recordset` | ✅ |
| `Scripting.FileSystemObject` + `TextStream` | ✅ |
| ASP intrinsic objects (Request, Response, Session, Server, Application) | ✅ |
| `Response.Write` (statement + expression) | ✅ |
| Line continuation (`_`) / `:` separator | ✅ |
| Date literals (`#...#`) | ✅ |
| HEX (`&HFF`) & octal (`&77`) literals | ✅ |
| Logical operators (`Eqv`, `Imp`) | ✅ |
| For loops (`Step`, negative `Step`) | ✅ |
| `Do` / `Loop` (While, Until, pre/post-test) | ✅ |
| `While` / `Wend` | ✅ |
| Comparison operators (`Is`, `=`, `<>`, etc.) | ✅ |
| Comments (`'`, `REM`) | ✅ |
| `<%@ %>` directives (LANGUAGE, ENABLESESSIONSTATE, CODEPAGE, LCID, TRANSACTION) | ✅ |
| `<!-- #include file="..." -->` / `virtual="..."` | ✅ |
| `Server.Execute` / `Server.Transfer` | ✅ |
| `Request.TotalBytes` | ✅ |
| Multipart form data | ✅ |
| `Application.Lock` / `.Unlock` (global mutex) | ✅ |
| VS Code DAP debugging | ✅ |

## Architecture

```
HTTP request
    │
    ▼
  ┌──────────────────┐
  │ IncludeResolver  │  Expands <!-- #include ... --> at source level
  └──────┬───────────┘
         │
  ┌──────▼───────────┐
  │  ASP Parser       │  Splits file into Html, Code, and Directive blocks
  └──────┬───────────┘
         │
  ┌──────▼───────────┐
  │  Preprocessor    │  Consumes Directive blocks → sets page configuration
  └──────┬───────────┘
         │
  ┌──────▼───────────┐
  │  Tokenizer        │  Lexes VBScript source into tokens
  └──────┬───────────┘
         │
  ┌──────▼───────────┐
  │   Parser          │  Pratt expression parser + recursive block parser
  └──────┬───────────┘
         │
  ┌──────▼────────────┐
  │  Interpreter       │  Walks the AST, evaluates expressions, executes blocks
  │  + Objects         │  ASP objects, COM objects, built-in functions
  └──────┬────────────┘
         │
  ┌──────▼───────────┐
  │  Response         │  Buffered output → HTTP 200 response
  └──────────────────┘
```

## Debugging

Full VS Code debugging support via Debug Adapter Protocol (DAP). See [`DEBUG.md`](DEBUG.md) for setup.

```bash
# Run the debug adapter standalone (launched by VS Code)
cargo run --bin asperger-debug -- --help
```

## Building from source

Requires [Rust](https://rustup.rs/) (edition 2021).

```bash
git clone <repo-url>
cd ASPerger
cargo build --release
./target/release/asperger --port 9090 --folder ./asp_files
```

## License

GNU General Public License v3.0
