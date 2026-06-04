# Debugging with ASPerger

ASPerger supports debugging ASP Classic / VBScript files through the Debug Adapter Protocol (DAP), allowing step-through debugging directly from VS Code.

## How it works

The debugger consists of three parts:

1. **`Debugger` struct** (`src/vbscript/debugger.rs`) — embedded in the interpreter; controls breakpoints, stepping, and call-stack tracking via an mpsc channel-based pause/resume mechanism
2. **`asperger-debug` binary** (`src/bin/asperger_debug.rs`) — a DAP server that reads requests from stdin and writes responses to stdout; spawns the interpreter in a separate thread and bridges DAP commands to the `Debugger` channel
3. **VS Code extension** (`extension/`) — registers `asperger-debug` as the debug adapter for `.asp` files

### Protocol flow

```
VS Code  <--DAP (stdin/stdout)-->  asperger-debug  <--mpsc-->  interpreter thread
```

- VS Code sends DAP requests (setBreakpoints, continue, next, etc.) to `asperger-debug`
- `asperger-debug` translates them into `DebugCommand` messages sent over a channel to the interpreter
- The interpreter calls `Debugger::check()` on each statement; when paused it blocks on `command_rx.recv()` waiting for the next command
- Events (stopped, output, exited) are sent back to VS Code via `server.output`

## Setup

### 1. Build the debug adapter

```bash
cargo build --bin asperger-debug
```

The binary is at `target/debug/asperger-debug`.

### 2. Install the VS Code extension

First, symlink the debug adapter binary into the extension directory:

```bash
mkdir -p extension/bin
ln -sf "$(pwd)/target/debug/asperger-debug" extension/bin/asperger-debug
```

Then symlink the extension folder into the VS Code extensions directory (must use `publisher.name` format):

```bash
# Linux/macOS
ln -sf "$(pwd)/extension" ~/.vscode/extensions/ssava.asperger-debug
```

Then reload VS Code (`Ctrl+Shift+P` → "Developer: Reload Window").

### 3. Create a launch configuration

In your `.vscode/launch.json`:

```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "type": "asperger",
            "request": "launch",
            "name": "Debug ASP file",
            "program": "${workspaceFolder}/path/to/file.asp"
        }
    ]
}
```

The `program` field points to the `.asp` file you want to debug.

## Features

### Breakpoints

- Set breakpoints by clicking the gutter in VS Code
- Breakpoints are checked on every statement — execution pauses at the exact line

### Stepping

| Action | Behavior |
|--------|----------|
| **Continue** (F5) | Resumes execution until the next breakpoint |
| **Step Over** (F10) | Executes the current line and pauses on the next |
| **Step Into** (F11) | Steps into function/sub calls |
| **Step Out** (Shift+F11) | Runs until the current function returns |

### Call stack

The call stack tracks `Function` and `Sub` calls. The top-level script appears as `Top Level`. Each frame shows the file, line number, and local variables.

### Variables

Local variables for the selected stack frame are displayed in the VS Code variables pane. Variable types are shown as `Double`, `String`, `Boolean`, `Null`, `Empty`, `Array`, or `Object`.

## Implementation notes

### Debugger hook

The single hook point in the interpreter is at `src/vbscript/block.rs:1520`:

```rust
if let Some(ref debugger) = context.debugger {
    let frame_depth = debugger.current_frame_depth();
    debugger.check("", 0, frame_depth)?;
}
```

Stack frames are pushed/popped in `execute_user_defined_function` (`src/vbscript/block.rs:1380-1406`).

### Channels

- `command_tx` / `command_rx`: `mpsc::Sender<DebugCommand>` / `Receiver<DebugCommand>` — DAP binary writes, interpreter reads
- `DebuggerState` is shared via `Arc<Mutex<DebuggerState>>` — breakpoints and stepping mode are written by the DAP thread and read by the interpreter thread

### Step modes

| Mode | Behavior |
|------|----------|
| `StepMode::Continue` | Only stop at breakpoints |
| `StepMode::StepOver` | Stop when frame depth <= saved depth or line changes |
| `StepMode::StepIn` | Stop at every statement |
| `StepMode::StepOut` | Stop when frame depth < saved depth |

### Limitations

- No expression evaluation (watch/hover) yet
- No conditional breakpoints
- Variables are read-only (no set-variable support)
- Single-threaded debugger (one thread ID)
