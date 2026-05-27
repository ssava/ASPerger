# ASPerger

A lightweight ASP Classic / VBScript server written in Rust.

## Features

- **ASP file parsing** — serves `.asp` files with `<% %>` code blocks
- **Full VBScript interpreter** — custom tokenizer, expression parser, and block evaluator
- **Control flow** — `If/Then/ElseIf/Else/End If`, `For/Next`, `For Each/Next`, `While/Wend`, `Do/Loop` (pre/post-test, While/Until)
- **Operators** — arithmetic (`+`, `-`, `*`, `/`, `\`, `^`, `Mod`), comparison (`=`, `<>`, `<`, `>`, `<=`, `>=`, `Is`), logical (`And`, `Or`, `Not`, `Xor`, `Eqv`, `Imp`), string concatenation (`&`, `+`)
- **Built-in functions** — `Len`, `UCase`, `LCase`, `Mid`, `Left`, `Right`, `Trim`, `CInt`, `CStr`, `Abs`, `IsNull`, `IsEmpty`, `InStr`, `Array`, `CreateObject`
- **Objects** — `Scripting.Dictionary` with `Count`/`Keys`/`Items` properties and `Add`/`Remove`/`Exists`/`RemoveAll` methods, indexed access (`dict(key)`)
- **Classes** — `Class`/`End Class` with `Public`/`Private` members, `Property Get`/`Property Let`
- **Error handling** — `On Error Resume Next` / `On Error Goto 0`, `Err.Number` / `Err.Description`
- **Response.Write** — output from VBScript blocks into the response
- **Static file serving** — CSS, JS, HTML, TXT, and other static files
- **Path traversal protection** — canonical path validation
- **Line continuation** — `_` for multi-line statements
- **Comments** — `'` and `REM`
- **Statement separator** — `:` (colon)

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

The repository includes a test suite at `asp_files/index.asp` that exercises most of the interpreter's capabilities. Run the server and open the page to see:

- `Response.Write` (with and without parentheses)
- Variable declaration (`Dim`) and assignment
- String and numeric variables
- `If/Then/ElseIf/Else/End If` conditions
- HTML output inside Response.Write
- Multi-part concatenated output
- `For i = 1 To 5 ... Next` loops
- `While ... Wend` loops
- User-defined `Function` / `Sub` with `Call`
- `Array` and `ReDim` / `ReDim Preserve`
- `Select Case`
- `Class` / `End Class` with `With` / `End With`
- `Property Get` / `Property Let`
- `On Error Resume Next` and `Err` object
- `Do ... Loop Until` (post-test)
- `For Each ... In ... Next` with `Scripting.Dictionary`
- Comparison operators (`=`, `Is`, `>=`, `<=`)
- String concatenation (`&` and `+`)
- `Mod` and integer division (`\`)
- `Empty`, `Null`, `Nothing` and `IsEmpty`/`IsNull` checks
- `Eqv` and `Imp` logical operators

## VBScript support

| Category | Status |
|----------|--------|
| Variables (`Dim`, `Set`, assignment) | ✅ |
| Control flow (If, For, While, Do, Select) | ✅ |
| Functions (`Function`, `Sub`, `Call`) | ✅ |
| Class (`Class`/`End Class`, Properties) | ✅ |
| Arrays (`Array()`, `ReDim`, index access) | ✅ |
| Built-in functions (string, numeric, type) | ✅ |
| `Scripting.Dictionary` | ✅ |
| `Response.Write` | ✅ |
| Error handling (`On Error Resume Next`) | ✅ |
| Line continuation (`_`) / `:` separator | ✅ |
| Date literals (`#...#`) | ✅ |
| `With` / `End With` | ✅ |
| HEX & octal literals (`&HFF`, `&77`) | ✅ |

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
