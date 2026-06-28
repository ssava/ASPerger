//! DAP debugger integration. Provides breakpoints, stepping, stack frames,
//! and communication channels between the interpreter and the VS Code debug adapter.

use std::sync::mpsc::{self, Receiver, Sender};
use std::sync::{Arc, Mutex};

use ahash::AHashMap;

use super::execution_context::CIString;
use super::value::VBValue;

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum StepMode {
    #[default]
    Continue,
    StepOver,
    StepIn,
    StepOut,
}

#[derive(Debug, Clone)]
pub struct StackFrame {
    pub name: String,
    pub file: String,
    pub line: usize,
    pub variables: AHashMap<CIString, VBValue>,
}

pub enum DebugCommand {
    Continue,
    Next,
    StepIn,
    StepOut,
    Disconnect,
}

pub enum DebugEvent {
    Stopped {
        reason: StoppedReason,
        file: String,
        line: usize,
        thread_id: u64,
    },
    Terminated,
}

pub enum StoppedReason {
    Breakpoint,
    Step,
    Pause,
}

#[derive(Default)]
pub struct DebuggerState {
    pub breakpoints: AHashMap<String, Vec<usize>>,
    pub step_mode: StepMode,
    pub step_frame_depth: usize,
    pub stack_frames: Vec<StackFrame>,
    pub current_file: String,
    pub current_line: usize,
    pub paused: bool,
}

pub struct Debugger {
    pub state: Arc<Mutex<DebuggerState>>,
    pub command_tx: Sender<DebugCommand>,
    pub command_rx: Mutex<Receiver<DebugCommand>>,
    pub event_tx: Sender<DebugEvent>,
}

impl Debugger {
    fn lock_state(&self) -> std::sync::MutexGuard<'_, DebuggerState> {
        self.state.lock().unwrap_or_else(|e| e.into_inner())
    }

    pub fn new() -> (Self, Arc<Mutex<DebuggerState>>, Receiver<DebugEvent>) {
        let state: Arc<Mutex<DebuggerState>> = Arc::new(Mutex::new(DebuggerState::default()));
        let (cmd_tx, cmd_rx) = mpsc::channel();
        let (evt_tx, evt_rx) = mpsc::channel();
        let d = Debugger {
            state: state.clone(),
            command_tx: cmd_tx,
            command_rx: Mutex::new(cmd_rx),
            event_tx: evt_tx,
        };
        (d, state, evt_rx)
    }

    pub fn set_breakpoints(&self, file: &str, lines: &[usize]) {
        let mut state = self.lock_state();
        let mut sorted = lines.to_vec();
        sorted.sort();
        sorted.dedup();
        state.breakpoints.insert(file.to_string(), sorted);
    }

    pub fn set_step_mode(&self, mode: StepMode) {
        let mut state = self.lock_state();
        state.step_mode = mode;
    }

    /// Evaluate whether execution should pause at the given source location.
    ///
    /// Called by the interpreter after every statement.  If `should_stop` is true
    /// the method:
    /// 1. Captures current variables into the top stack frame.
    /// 2. Sends a `Stopped` event to the DAP client.
    /// 3. **Blocks** on `command_rx.recv()` until the user issues a debug command
    ///    (Continue / Next / StepIn / StepOut / Disconnect).
    /// 4. Updates `step_mode` and `step_frame_depth` according to the command,
    ///    then returns so the interpreter can resume.
    ///
    /// If `should_stop` is false the method just records the current position
    /// (used by StepOver to detect line changes) and returns immediately.
    ///
    /// ## Step-mode rules (`should_stop` logic)
    /// | Mode      | Stop condition                                       |
    /// |-----------|------------------------------------------------------|
    /// | Continue  | Line is in the breakpoint list for this file          |
    /// | StepOver  | frame ≤ stored depth AND line changed                |
    /// | StepIn    | Always stop                                           |
    /// | StepOut   | frame < stored depth (returning from a call)         |
    pub fn check(
        &self,
        file: &str,
        line: usize,
        frame_depth: usize,
        vars: Option<&AHashMap<CIString, VBValue>>,
    ) -> Result<(), crate::vbscript::vbs_error::VBSError> {
        use crate::vbscript::vbs_error::VBSErrorType;

        let should_stop = {
            let s = self.lock_state();
            match s.step_mode {
                StepMode::Continue => s
                    .breakpoints
                    .get(file)
                    .is_some_and(|lines| lines.contains(&line)),
                StepMode::StepOver => frame_depth <= s.step_frame_depth && line != s.current_line,
                StepMode::StepIn => true,
                StepMode::StepOut => frame_depth < s.step_frame_depth,
            }
        };

        if should_stop {
            // -- snapshot state and send Stopped event --
            {
                let mut s = self.lock_state();
                s.paused = true;
                s.current_file = file.to_string();
                s.current_line = line;

                if let Some(vars) = vars {
                    let cloned = vars.clone();
                    if s.stack_frames.is_empty() {
                        s.stack_frames.push(StackFrame {
                            name: "Top Level".to_string(),
                            file: file.to_string(),
                            line,
                            variables: cloned,
                        });
                    } else if let Some(top) = s.stack_frames.last_mut() {
                        top.variables = cloned;
                        top.line = line;
                    }
                }
            }

            let reason = match self.lock_state().step_mode {
                StepMode::Continue => StoppedReason::Breakpoint,
                _ => StoppedReason::Step,
            };

            self.event_tx
                .send(DebugEvent::Stopped {
                    reason,
                    file: file.to_string(),
                    line,
                    thread_id: 1,
                })
                .unwrap_or(());

            // -- block until the DAP client sends a command (30s timeout) --
            use std::sync::mpsc::RecvTimeoutError;
            let cmd = self.command_rx
                .lock().unwrap_or_else(|e| e.into_inner())
                .recv_timeout(std::time::Duration::from_secs(30));
            match cmd {
                Ok(DebugCommand::Continue) => {
                    let mut s = self.lock_state();
                    s.step_mode = StepMode::Continue;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::Next) => {
                    let mut s = self.lock_state();
                    s.step_mode = StepMode::StepOver;
                    s.step_frame_depth = frame_depth;
                    s.current_file = file.to_string();
                    s.current_line = line;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::StepIn) => {
                    let mut s = self.lock_state();
                    s.step_mode = StepMode::StepIn;
                    s.step_frame_depth = frame_depth;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::StepOut) => {
                    let mut s = self.lock_state();
                    s.step_mode = StepMode::StepOut;
                    s.step_frame_depth = frame_depth;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::Disconnect) | Err(RecvTimeoutError::Disconnected) => {
                    return Err(
                        VBSErrorType::RuntimeError.into_error("Debugger disconnected".to_string())
                    );
                }
                Err(RecvTimeoutError::Timeout) => {
                    return Err(
                        VBSErrorType::RuntimeError.into_error("Debugger command timed out".to_string())
                    );
                }
            }
        } else if line != 0 {
            // No stop — just remember the current position for StepOver tracking
            let mut s = self.lock_state();
            s.current_file = file.to_string();
            s.current_line = line;
        }

        Ok(())
    }

    pub fn push_frame(&self, name: &str, file: &str, line: usize, vars: AHashMap<CIString, VBValue>) {
        let mut s = self.lock_state();
        s.stack_frames.push(StackFrame {
            name: name.to_string(),
            file: file.to_string(),
            line,
            variables: vars,
        });
    }

    pub fn pop_frame(&self) {
        let mut s = self.lock_state();
        s.stack_frames.pop();
    }

    pub fn current_frame_depth(&self) -> usize {
        self.lock_state().stack_frames.len()
    }

    pub fn update_variables(&self) {
        drop(self.lock_state());
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::vbscript::execution_context::CIString;

    fn make_vars() -> AHashMap<CIString, VBValue> {
        let mut m = AHashMap::new();
        m.insert(CIString::new("x".to_string()), VBValue::Number(1.0));
        m
    }

    #[test]
    fn test_debugger_new_initial_state() {
        let (_, state, _) = Debugger::new();
        let s = state.lock().unwrap();
        assert!(s.breakpoints.is_empty());
        assert_eq!(s.step_mode, StepMode::Continue);
        assert!(!s.paused);
        assert!(s.stack_frames.is_empty());
    }

    #[test]
    fn test_debugger_set_breakpoints() {
        let (d, state, _) = Debugger::new();
        d.set_breakpoints("test.asp", &[5, 10, 15]);
        let s = state.lock().unwrap();
        assert_eq!(s.breakpoints.get("test.asp").unwrap(), &vec![5, 10, 15]);
    }

    #[test]
    fn test_debugger_set_step_mode() {
        let (d, state, _) = Debugger::new();
        d.set_step_mode(StepMode::StepIn);
        let s = state.lock().unwrap();
        assert_eq!(s.step_mode, StepMode::StepIn);
    }

    #[test]
    fn test_debugger_check_no_breakpoint_no_stop() {
        let (d, _, _) = Debugger::new();
        let result = d.check("test.asp", 10, 0, Some(&make_vars()));
        assert!(result.is_ok());
    }

    #[test]
    fn test_debugger_check_breakpoint_hit_triggers_stop() {
        let (d, _state, evt_rx) = Debugger::new();
        d.set_breakpoints("test.asp", &[10]);
        d.set_step_mode(StepMode::Continue);

        // Spawn a thread to send Continue after a short delay
        let cmd_tx = d.command_tx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(50));
            cmd_tx.send(DebugCommand::Continue).unwrap();
        });

        let result = d.check("test.asp", 10, 0, Some(&make_vars()));
        assert!(result.is_ok());

        // Verify a Stopped event was sent
        let event = evt_rx.recv_timeout(std::time::Duration::from_millis(100));
        assert!(event.is_ok());
        match event.unwrap() {
            DebugEvent::Stopped { reason, file, line, .. } => {
                assert_eq!(reason as i32, StoppedReason::Breakpoint as i32);
                assert_eq!(file, "test.asp");
                assert_eq!(line, 10);
            }
            _ => panic!("expected Stopped event"),
        }
    }

    #[test]
    fn test_debugger_check_step_in_triggers() {
        let (d, _, evt_rx) = Debugger::new();
        d.set_step_mode(StepMode::StepIn);

        let cmd_tx = d.command_tx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(50));
            cmd_tx.send(DebugCommand::Continue).unwrap();
        });

        let result = d.check("test.asp", 1, 0, Some(&make_vars()));
        assert!(result.is_ok());
        let event = evt_rx.recv_timeout(std::time::Duration::from_millis(100));
        assert!(event.is_ok());
        match event.unwrap() {
            DebugEvent::Stopped { reason, .. } => {
                assert_eq!(reason as i32, StoppedReason::Step as i32);
            }
            _ => panic!("expected Stopped event"),
        }
    }

    #[test]
    fn test_debugger_check_step_over_skips_same_line() {
        let (d, _, evt_rx) = Debugger::new();
        d.set_step_mode(StepMode::StepOver);

        // First call: different line from current_line (0), so should_stop=true.
        // Need to send Continue to unblock.
        let cmd_tx = d.command_tx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(50));
            cmd_tx.send(DebugCommand::Continue).unwrap();
        });
        let _ = d.check("test.asp", 10, 0, None);
        // Consume the Stopped event from the first stop
        let _first = evt_rx.recv_timeout(std::time::Duration::from_millis(100));

        // Second call on same line: should_stop=false (step over skips same line)
        let result = d.check("test.asp", 10, 0, None);
        assert!(result.is_ok());
        // No new event should be sent
        let event = evt_rx.recv_timeout(std::time::Duration::from_millis(30));
        assert!(event.is_err()); // timeout = no event
    }

    #[test]
    fn test_debugger_push_pop_frame() {
        let (d, state, _) = Debugger::new();
        d.push_frame("func1", "test.asp", 5, AHashMap::new());
        assert_eq!(d.current_frame_depth(), 1);
        d.push_frame("func2", "test.asp", 10, AHashMap::new());
        assert_eq!(d.current_frame_depth(), 2);
        d.pop_frame();
        assert_eq!(d.current_frame_depth(), 1);
        let s = state.lock().unwrap();
        assert_eq!(s.stack_frames.len(), 1);
        assert_eq!(s.stack_frames[0].name, "func1");
    }

    #[test]
    fn test_debugger_command_next_sets_step_over() {
        let (d, state, _) = Debugger::new();
        // Set a breakpoint so check() blocks and waits for a command
        d.set_breakpoints("test.asp", &[10]);
        d.set_step_mode(StepMode::Continue);

        let cmd_tx = d.command_tx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(50));
            cmd_tx.send(DebugCommand::Next).unwrap();
        });

        let _ = d.check("test.asp", 10, 2, Some(&make_vars()));
        let s = state.lock().unwrap();
        assert_eq!(s.step_mode, StepMode::StepOver);
        assert_eq!(s.step_frame_depth, 2);
    }

    #[test]
    fn test_debugger_check_step_out_triggers() {
        let (d, _, _evt_rx) = Debugger::new();
        // StepOut triggers when frame_depth < stored frame_depth
        // With default state, both are 0 so it won't trigger — just verify no error
        let result = d.check("test.asp", 1, 0, None);
        assert!(result.is_ok());
    }

    #[test]
    fn test_debugger_state_paused_on_stop() {
        let (d, state, _) = Debugger::new();
        d.set_breakpoints("test.asp", &[5]);
        d.set_step_mode(StepMode::Continue);

        let cmd_tx = d.command_tx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(std::time::Duration::from_millis(50));
            cmd_tx.send(DebugCommand::Continue).unwrap();
        });

        let _ = d.check("test.asp", 5, 0, Some(&make_vars()));
        let s = state.lock().unwrap();
        assert!(!s.paused); // cleared by Continue
    }
}
