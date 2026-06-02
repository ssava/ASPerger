use std::io::{BufReader, BufWriter, Write};
use std::sync::{Arc, Mutex};

use dap::events;
use dap::prelude::*;
use dap::responses;
use dap::types;

use asperger::vbscript::debugger::{DebugCommand, DebugEvent, Debugger, StoppedReason};
use asperger::vbscript::execution_context::ExecutionContext;
use asperger::vbscript::interpreter::VBScriptInterpreter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let stdin = std::io::stdin();
    let stdout = std::io::stdout();

    let reader = BufReader::new(stdin.lock());
    let writer = BufWriter::new(stdout);

    let mut server = Server::new(reader, writer);
    let server_output = server.output.clone();

    let mut script_path = String::new();
    let mut debugger_state: Option<Arc<Mutex<asperger::vbscript::debugger::DebuggerState>>> = None;
    let mut command_tx: Option<std::sync::mpsc::Sender<DebugCommand>> = None;
    let mut interpreter_handle: Option<std::thread::JoinHandle<()>> = None;

    loop {
        let req = match server.poll_request()? {
            Some(r) => r,
            None => return Ok(()),
        };

        match req.command {
            Command::Initialize(_) => {
                let caps = types::Capabilities {
                    supports_configuration_done_request: Some(true),
                    supports_function_breakpoints: Some(false),
                    supports_conditional_breakpoints: Some(false),
                    supports_hit_conditional_breakpoints: Some(false),
                    supports_evaluate_for_hovers: Some(false),
                    supports_step_back: Some(false),
                    supports_set_variable: Some(false),
                    supports_restart_frame: Some(false),
                    supports_goto_targets_request: Some(false),
                    supports_step_in_targets_request: Some(false),
                    supports_completions_request: Some(false),
                    supports_modules_request: Some(false),
                    supports_restart_request: Some(false),
                    supports_exception_options: Some(false),
                    supports_value_formatting_options: Some(false),
                    supports_exception_info_request: Some(false),
                    support_terminate_debuggee: Some(false),
                    support_suspend_debuggee: Some(false),
                    supports_delayed_stack_trace_loading: Some(false),
                    supports_loaded_sources_request: Some(false),
                    supports_log_points: Some(false),
                    supports_terminate_threads_request: Some(false),
                    supports_set_expression: Some(false),
                    supports_terminate_request: Some(false),
                    supports_data_breakpoints: Some(false),
                    supports_read_memory_request: Some(false),
                    supports_write_memory_request: Some(false),
                    supports_disassemble_request: Some(false),
                    supports_cancel_request: Some(false),
                    supports_stepping_granularity: Some(false),
                    supports_instruction_breakpoints: Some(false),
                    supports_exception_filter_options: Some(false),
                    supports_single_thread_execution_requests: Some(false),
                    ..Default::default()
                };
                let rsp = req.success(ResponseBody::Initialize(caps));
                server.respond(rsp)?;
                server.send_event(Event::Initialized)?;
            }

            Command::Launch(ref args) => {
                if let Some(ref extra) = args.additional_data {
                    if let Some(program) = extra.get("program").and_then(|v| v.as_str()) {
                        script_path = program.to_string();
                    }
                }
                let rsp = req.success(ResponseBody::Launch);
                server.respond(rsp)?;
            }

            Command::SetBreakpoints(ref args) => {
                let path = args.source.path.clone().unwrap_or_default();
                let lines: Vec<usize> = args
                    .breakpoints
                    .as_ref()
                    .map(|bps| {
                        bps.iter()
                            .map(|bp| bp.line as usize)
                            .collect()
                    })
                    .unwrap_or_default();

                if let Some(ref state) = debugger_state {
                    let mut s = state.lock().unwrap();
                    s.breakpoints.insert(path.clone(), lines.clone());
                }

                let bps_response: Vec<types::Breakpoint> = lines
                    .iter()
                    .map(|&l| types::Breakpoint {
                        id: Some(l as i64),
                        verified: true,
                        line: Some(l as i64),
                        source: Some(types::Source {
                            path: Some(path.clone()),
                            name: None,
                            source_reference: None,
                            presentation_hint: None,
                            origin: None,
                            sources: None,
                            adapter_data: None,
                            checksums: None,
                        }),
                        column: None,
                        end_line: None,
                        end_column: None,
                        message: None,
                        instruction_reference: None,
                        offset: None,
                        })
                        .collect();
                let rsp = req.success(ResponseBody::SetBreakpoints(
                    responses::SetBreakpointsResponse {
                        breakpoints: bps_response,
                    },
                ));
                server.respond(rsp)?;
            }

            Command::ConfigurationDone => {
                let rsp = req.success(ResponseBody::ConfigurationDone);
                server.respond(rsp)?;

                let (debugger, state, event_rx) = Debugger::new();
                debugger_state = Some(state.clone());
                command_tx = Some(debugger.command_tx.clone());

                // Forward DebugEvents from interpreter thread to VS Code
                let evt_so = server_output.clone();
                std::thread::spawn(move || {
                    while let Ok(event) = event_rx.recv() {
                        let mut so = match evt_so.lock() {
                            Ok(s) => s,
                            Err(_) => break,
                        };
                        match event {
                            DebugEvent::Stopped { reason, thread_id, .. } => {
                                let dap_reason = match reason {
                                    StoppedReason::Breakpoint => types::StoppedEventReason::Breakpoint,
                                    StoppedReason::Step => types::StoppedEventReason::Step,
                                    StoppedReason::Pause => types::StoppedEventReason::Pause,
                                };
                                let _ = so.send_event(Event::Stopped(events::StoppedEventBody {
                                    reason: dap_reason,
                                    description: None,
                                    thread_id: Some(thread_id as i64),
                                    preserve_focus_hint: None,
                                    text: None,
                                    all_threads_stopped: Some(true),
                                    hit_breakpoint_ids: None,
                                }));
                            }
                            DebugEvent::Terminated => break,
                        }
                    }
                });

                let script = script_path.clone();
                let so = server_output.clone();

                let handle = std::thread::spawn(move || {
                    run_debug_session(&script, debugger, so);
                });
                interpreter_handle = Some(handle);

                // Set initial step mode to step-in so we stop at the first line
                {
                    let mut s = state.lock().unwrap();
                    s.step_mode = asperger::vbscript::debugger::StepMode::StepIn;
                }
            }

            Command::Continue(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::Continue).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::Continue(responses::ContinueResponse {
                    all_threads_continued: Some(true),
                }));
                server.respond(rsp)?;
            }

            Command::Next(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::Next).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::Next);
                server.respond(rsp)?;
            }

            Command::StepIn(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::StepIn).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::StepIn);
                server.respond(rsp)?;
            }

            Command::StepOut(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::StepOut).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::StepOut);
                server.respond(rsp)?;
            }

            Command::Pause(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::Continue).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::Pause);
                server.respond(rsp)?;
            }

            Command::StackTrace(_) => {
                let frames = if let Some(ref state) = debugger_state {
                    let s = state.lock().unwrap();
                    let mut frames: Vec<types::StackFrame> = s
                        .stack_frames
                        .iter()
                        .enumerate()
                        .map(|(i, f)| types::StackFrame {
                            id: i as i64,
                            name: f.name.clone(),
                            source: Some(types::Source {
                                path: Some(f.file.clone()),
                                name: None,
                                source_reference: None,
                                presentation_hint: None,
                                origin: None,
                                sources: None,
                                adapter_data: None,
                                checksums: None,
                            }),
                            line: f.line as i64,
                            column: 1,
                            end_line: None,
                            end_column: None,
                            can_restart: Some(false),
                            instruction_pointer_reference: None,
                            module_id: None,
                            presentation_hint: None,
                        })
                        .collect();
                    frames.push(types::StackFrame {
                        id: s.stack_frames.len() as i64,
                        name: "Top Level".to_string(),
                        source: Some(types::Source {
                            path: Some(s.current_file.clone()),
                            name: None,
                            source_reference: None,
                            presentation_hint: None,
                            origin: None,
                            sources: None,
                            adapter_data: None,
                            checksums: None,
                        }),
                        line: s.current_line as i64,
                        column: 1,
                        end_line: None,
                        end_column: None,
                        can_restart: Some(false),
                        instruction_pointer_reference: None,
                        module_id: None,
                        presentation_hint: None,
                    });
                    frames
                } else {
                    vec![]
                };
                let rsp = req.success(ResponseBody::StackTrace(responses::StackTraceResponse {
                    stack_frames: frames,
                    total_frames: None,
                }));
                server.respond(rsp)?;
            }

            Command::Scopes(ref args) => {
                let frame_id = args.frame_id;
                let scopes = vec![types::Scope {
                    name: "Locals".to_string(),
                    presentation_hint: Some(types::ScopePresentationhint::Locals),
                    variables_reference: frame_id + 1000,
                    named_variables: None,
                    indexed_variables: None,
                    expensive: false,
                    source: None,
                    line: None,
                    column: None,
                    end_line: None,
                    end_column: None,
                }];
                let rsp = req.success(ResponseBody::Scopes(responses::ScopesResponse { scopes }));
                server.respond(rsp)?;
            }

            Command::Variables(ref args) => {
                let ref_id = args.variables_reference;

                let variables: Vec<types::Variable> = if let Some(ref state) = debugger_state {
                    let s = state.lock().unwrap();
                    let frame_idx = (ref_id.saturating_sub(1000)) as usize;
                    let vars = s
                        .stack_frames
                        .get(frame_idx)
                        .map(|f| &f.variables)
                        .unwrap();
                    vars.iter()
                        .map(|(k, v)| types::Variable {
                            name: k.clone(),
                            value: format!("{:?}", v),
                            type_field: Some(match v {
                                asperger::vbscript::VBValue::Number(_) => "Double".to_string(),
                                asperger::vbscript::VBValue::String(_) => "String".to_string(),
                                asperger::vbscript::VBValue::Boolean(_) => "Boolean".to_string(),
                                asperger::vbscript::VBValue::Null => "Null".to_string(),
                                asperger::vbscript::VBValue::Empty => "Empty".to_string(),
                                asperger::vbscript::VBValue::Array(_) => "Array".to_string(),
                                asperger::vbscript::VBValue::Object(_) => "Object".to_string(),
                            }),
                            variables_reference: 0,
                            named_variables: None,
                            indexed_variables: None,
                            evaluate_name: None,
                            memory_reference: None,
                            presentation_hint: None,
                        })
                        .collect()
                } else {
                    vec![]
                };
                let rsp = req.success(ResponseBody::Variables(responses::VariablesResponse {
                    variables,
                }));
                server.respond(rsp)?;
            }

            Command::Disconnect(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::Disconnect).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::Disconnect);
                server.respond(rsp)?;
                if let Some(handle) = interpreter_handle.take() {
                    let _ = handle.join();
                }
                break;
            }

            _ => {
                let rsp = req.success(ResponseBody::ConfigurationDone);
                server.respond(rsp)?;
            }
        }
    }

    Ok(())
}

fn run_debug_session<W: Write + Send + 'static>(
    script_path: &str,
    debugger: Debugger,
    server_output: Arc<Mutex<dap::server::ServerOutput<W>>>,
) {
    let script = match std::fs::read_to_string(script_path) {
        Ok(c) => c,
        Err(e) => {
            let mut so = server_output.lock().unwrap();
            let _ = so.send_event(Event::Output(events::OutputEventBody {
                output: format!("Cannot read script '{}': {}\n", script_path, e),
                category: Some(types::OutputEventCategory::Stderr),
                ..Default::default()
            }));
            return;
        }
    };

    let mut ctx = ExecutionContext::new();
    ctx.debugger = Some(debugger);

    let interp = VBScriptInterpreter;
    let result = interp.execute(&script, &mut ctx);

    let is_ok = result.is_ok();
    {
        let mut so = server_output.lock().unwrap();
        if let Err(e) = result {
            let _ = so.send_event(Event::Output(events::OutputEventBody {
                output: format!("Runtime error: {}\n", e),
                category: Some(types::OutputEventCategory::Stderr),
                ..Default::default()
            }));
        }
        let _ = so.send_event(Event::Exited(events::ExitedEventBody {
            exit_code: if is_ok { 0 } else { 1 },
        }));
        let _ = so.send_event(Event::Terminated(None));
    }
}
