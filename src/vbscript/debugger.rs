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
        let mut state = self.state.lock().unwrap();
        let mut sorted = lines.to_vec();
        sorted.sort();
        sorted.dedup();
        state.breakpoints.insert(file.to_string(), sorted);
    }

    pub fn set_step_mode(&self, mode: StepMode) {
        let mut state = self.state.lock().unwrap();
        state.step_mode = mode;
    }

    pub fn check(
        &self,
        file: &str,
        line: usize,
        frame_depth: usize,
        vars: Option<&AHashMap<CIString, VBValue>>,
    ) -> Result<(), crate::vbscript::vbs_error::VBSError> {
        use crate::vbscript::vbs_error::VBSErrorType;

        let should_stop = {
            let s = self.state.lock().unwrap();
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
            {
                let mut s = self.state.lock().unwrap();
                s.paused = true;
                s.current_file = file.to_string();
                s.current_line = line;

                // Capture variables into the current frame
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

            let reason = match self.state.lock().unwrap().step_mode {
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

            // Wait for a debugger command (blocks until received)
            match self.command_rx.lock().unwrap().recv() {
                Ok(DebugCommand::Continue) => {
                    let mut s = self.state.lock().unwrap();
                    s.step_mode = StepMode::Continue;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::Next) => {
                    let mut s = self.state.lock().unwrap();
                    s.step_mode = StepMode::StepOver;
                    s.step_frame_depth = frame_depth;
                    s.current_file = file.to_string();
                    s.current_line = line;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::StepIn) => {
                    let mut s = self.state.lock().unwrap();
                    s.step_mode = StepMode::StepIn;
                    s.step_frame_depth = frame_depth;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::StepOut) => {
                    let mut s = self.state.lock().unwrap();
                    s.step_mode = StepMode::StepOut;
                    s.step_frame_depth = frame_depth;
                    s.paused = false;
                    s.stack_frames.clear();
                }
                Ok(DebugCommand::Disconnect) => {
                    return Err(
                        VBSErrorType::RuntimeError.into_error("Debugger disconnected".to_string())
                    );
                }
                Err(_) => {}
            }
        } else if line != 0 {
            let mut s = self.state.lock().unwrap();
            s.current_file = file.to_string();
            s.current_line = line;
        }

        Ok(())
    }

    pub fn push_frame(&self, name: &str, file: &str, line: usize, vars: AHashMap<CIString, VBValue>) {
        let mut s = self.state.lock().unwrap();
        s.stack_frames.push(StackFrame {
            name: name.to_string(),
            file: file.to_string(),
            line,
            variables: vars,
        });
    }

    pub fn pop_frame(&self) {
        let mut s = self.state.lock().unwrap();
        s.stack_frames.pop();
    }

    pub fn current_frame_depth(&self) -> usize {
        self.state.lock().unwrap().stack_frames.len()
    }

    pub fn update_variables(&self) {
        drop(self.state.lock().unwrap());
    }
}
