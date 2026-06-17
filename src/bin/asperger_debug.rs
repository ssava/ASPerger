use std::io::{BufReader, BufWriter, Write};
use std::sync::{Arc, Mutex};

use clap::Parser;
use dap::events;
use dap::prelude::*;
use dap::responses;
use dap::types;
use asperger::asp::config::Config as AspConfig;
use asperger::asp::server::AspServer;
use asperger::vbscript::debugger::{DebugCommand, DebugEvent, Debugger, StoppedReason};

#[derive(Parser)]
#[command(
    name = "asperger-debug",
    version,
    about = "VBScript debug adapter (DAP)"
)]
struct Cli {
    /// Path to the script to debug
    program: Option<String>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let cli = Cli::parse();

    let stdin = std::io::stdin();
    let stdout = std::io::stdout();

    let reader = BufReader::new(stdin.lock());
    let writer = BufWriter::new(stdout);

    let mut server = Server::new(reader, writer);
    let server_output = server.output.clone();

    let mut script_path = String::new();
    let mut launch_folder: Option<String> = None;
    let mut launch_port: Option<u16> = None;
    let mut launch_default_doc: Option<String> = None;
    let mut launch_directory_listing: Option<bool> = None;
    let mut debugger_state: Option<Arc<Mutex<asperger::vbscript::debugger::DebuggerState>>> = None;
    let mut command_tx: Option<std::sync::mpsc::Sender<DebugCommand>> = None;
    let mut interpreter_handle: Option<std::thread::JoinHandle<()>> = None;
    let mut pending_breakpoints: Vec<(String, Vec<usize>)> = Vec::new();
    let mut debugger_instance: Option<Arc<Debugger>> = None;

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
                    supports_evaluate_for_hovers: Some(true),
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
                    if let Some(folder) = extra.get("folder").and_then(|v| v.as_str()) {
                        if !folder.is_empty() {
                            launch_folder = Some(folder.to_string());
                        }
                    }
                    if let Some(port) = extra.get("port").and_then(|v| v.as_u64()) {
                        launch_port = Some(port as u16);
                    }
                    if let Some(doc) = extra.get("defaultDocument").and_then(|v| v.as_str()) {
                        if !doc.is_empty() {
                            launch_default_doc = Some(doc.to_string());
                        }
                    }
                    if let Some(dl) = extra.get("directoryListing").and_then(|v| v.as_bool()) {
                        launch_directory_listing = Some(dl);
                    }
                }
                if script_path.is_empty() {
                    if let Some(ref program) = cli.program {
                        script_path = program.to_string();
                    }
                }

                // Create Debugger early so SetBreakpoints (which arrives before
                // ConfigurationDone in DAP protocol) can store breakpoints
                let (debugger, state, event_rx) = Debugger::new();
                debugger_state = Some(state.clone());
                command_tx = Some(debugger.command_tx.clone());
                debugger_instance = Some(Arc::new(debugger));

                // Flush any breakpoints received before Launch
                if let Some(ref state) = debugger_state {
                    let mut s = state.lock().unwrap();
                    for (path, lines) in pending_breakpoints.drain(..) {
                        s.breakpoints.insert(path, lines);
                    }
                }

                // Forward DebugEvents from interpreter thread to VS Code
                let evt_so = server_output.clone();
                std::thread::spawn(move || {
                    while let Ok(event) = event_rx.recv() {
                        let mut so = match evt_so.lock() {
                            Ok(s) => s,
                            Err(_) => break,
                        };
                        match event {
                            DebugEvent::Stopped {
                                reason, thread_id, ..
                            } => {
                                let dap_reason = match reason {
                                    StoppedReason::Breakpoint => {
                                        types::StoppedEventReason::Breakpoint
                                    }
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

                let rsp = req.success(ResponseBody::Launch);
                server.respond(rsp)?;
            }

            Command::SetBreakpoints(ref args) => {
                let path = args.source.path.clone().unwrap_or_default();
                let lines: Vec<usize> = args
                    .breakpoints
                    .as_ref()
                    .map(|bps| bps.iter().map(|bp| bp.line as usize).collect())
                    .unwrap_or_default();

                if let Some(ref state) = debugger_state {
                    let mut s = state.lock().unwrap();
                    s.breakpoints.insert(path.clone(), lines.clone());
                } else {
                    // Buffer breakpoints until Debugger exists (after Launch)
                    pending_breakpoints.push((path.clone(), lines.clone()));
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

                let so = server_output.clone();
                let debugger = debugger_instance.take().unwrap();

                // Start in Continue mode — user must set explicit breakpoints
                if let Some(ref state) = debugger_state {
                    let mut s = state.lock().unwrap();
                    s.step_mode = asperger::vbscript::debugger::StepMode::Continue;
                }

                // Determine served root folder.
                let folder = launch_folder.clone().unwrap_or_else(|| {
                    let p = std::path::Path::new(&script_path);
                    if p.is_dir() {
                        p.to_string_lossy().to_string()
                    } else {
                        p.parent()
                            .map(|p| p.to_string_lossy().to_string())
                            .unwrap_or_else(|| ".".to_string())
                    }
                });

                // Build config: start with INI from folder, then apply launch overrides
                let mut config = asperger::asp::config::AspServerConfig::from_folder(&folder);
                config.apply_overrides(
                    None,                           // host — not settable from DAP yet
                    launch_port,
                    Some(&folder),                   // folder (always set)
                    launch_default_doc.as_deref(),
                    launch_directory_listing,
                );

                let handle = std::thread::spawn(move || {
                    start_debug_http_server(&config, debugger, so);
                });
                interpreter_handle = Some(handle);
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
                    if let Some(frame) = s.stack_frames.get(frame_idx) {
                        frame
                            .variables
                            .iter()
                            .map(|(k, v)| types::Variable {
                                name: k.clone(),
                                value: v.to_string(),
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
                    }
                } else {
                    vec![]
                };
                let rsp = req.success(ResponseBody::Variables(responses::VariablesResponse {
                    variables,
                }));
                server.respond(rsp)?;
            }

            Command::Evaluate(ref args) => {
                let expression = args.expression.trim().to_string();
                let frame_id = args.frame_id.unwrap_or(0);

                let (result, type_field) = if let Some(ref state) = debugger_state {
                    let s = state.lock().unwrap();
                    let upper = expression.to_uppercase();

                    // Look up the expression in the specified frame, then fall back to all frames
                    let found = s.stack_frames.get(frame_id as usize)
                        .and_then(|f| f.variables.get(&upper))
                        .or_else(|| {
                            s.stack_frames.iter().rev().find_map(|f| f.variables.get(&upper))
                        });

                    match found {
                        Some(v) => (v.to_string(), Some(match v {
                            asperger::vbscript::VBValue::Number(_) => "Double".to_string(),
                            asperger::vbscript::VBValue::String(_) => "String".to_string(),
                            asperger::vbscript::VBValue::Boolean(_) => "Boolean".to_string(),
                            asperger::vbscript::VBValue::Null => "Null".to_string(),
                            asperger::vbscript::VBValue::Empty => "Empty".to_string(),
                            asperger::vbscript::VBValue::Array(_) => "Array".to_string(),
                            asperger::vbscript::VBValue::Object(_) => "Object".to_string(),
                        })),
                        None => {
                            // Try to interpret as a literal
                            let trimmed = expression.trim();
                            if trimmed.eq_ignore_ascii_case("true") {
                                ("True".to_string(), Some("Boolean".to_string()))
                            } else if trimmed.eq_ignore_ascii_case("false") {
                                ("False".to_string(), Some("Boolean".to_string()))
                            } else if trimmed.parse::<f64>().is_ok() {
                                (trimmed.to_string(), Some("Double".to_string()))
                            } else if trimmed.starts_with('"') && trimmed.ends_with('"') && trimmed.len() >= 2 {
                                (trimmed[1..trimmed.len()-1].to_string(), Some("String".to_string()))
                            } else {
                                (format!("<error: '{}' not found>", expression), Some("Error".to_string()))
                            }
                        }
                    }
                } else {
                    ("<error: no debugger state>".to_string(), Some("Error".to_string()))
                };

                let rsp = req.success(ResponseBody::Evaluate(responses::EvaluateResponse {
                    result,
                    type_field,
                    presentation_hint: None,
                    variables_reference: 0,
                    named_variables: None,
                    indexed_variables: None,
                    memory_reference: None,
                }));
                server.respond(rsp)?;
            }

            Command::Threads => {
                let rsp = req.success(ResponseBody::Threads(responses::ThreadsResponse {
                    threads: vec![types::Thread {
                        id: 1,
                        name: "Main Thread".to_string(),
                    }],
                }));
                server.respond(rsp)?;
            }

            Command::Disconnect(_) => {
                if let Some(ref tx) = command_tx {
                    tx.send(DebugCommand::Disconnect).unwrap_or(());
                }
                let rsp = req.success(ResponseBody::Disconnect);
                server.respond(rsp)?;
                // Do NOT join the HTTP thread — it loops infinitely
                // The process exits after break, killing all threads
                if let Some(handle) = interpreter_handle.take() {
                    drop(handle); // detach, don't join
                }
                break;
            }

            _ => {
                // Respond with a generic success for unrecognized requests
                // to keep VS Code happy
                let rsp = req.success(ResponseBody::ConfigurationDone);
                server.respond(rsp)?;
            }
        }
    }

    Ok(())
}

fn start_debug_http_server<W: Write + Send + 'static>(
    config: &asperger::asp::config::AspServerConfig,
    debugger: Arc<Debugger>,
    server_output: Arc<Mutex<dap::server::ServerOutput<W>>>,
) {
    let bind_addr = format!("{}:{}", config.host, config.port);
    let folder = config.folder.clone();
    let default_document = config.default_document.clone();

    let asp_cfg = AspConfig {
        host: config.host.clone(),
        port: config.port,
        folder: folder.clone(),
        program: None,
        enable_directory_listing: config.directory_listing,
    };
    let server = AspServer::new(asp_cfg);

    // Use a single-thread Tokio runtime for all async I/O
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_io()
        .enable_time()
        .build()
        .unwrap();

    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        rt.block_on(async {
            let listener = match tokio::net::TcpListener::bind(&bind_addr).await {
                Ok(l) => l,
                Err(e) => {
                    let mut so = server_output.lock().unwrap();
                    let _ = so.send_event(Event::Output(events::OutputEventBody {
                        output: format!("Failed to bind debug server on {}: {}\n", bind_addr, e),
                        category: Some(types::OutputEventCategory::Stderr),
                        ..Default::default()
                    }));
                    return;
                }
            };

            {
                let mut so = server_output.lock().unwrap();
                let _ = so.send_event(Event::Output(events::OutputEventBody {
                    output: format!(
                        "ASP Debug Server started on http://{}:{}/ (folder={:?}, default={:?})\n",
                        config.host, config.port, folder, default_document
                    ),
                    category: Some(types::OutputEventCategory::Stdout),
                    ..Default::default()
                }));
            }

            loop {
                let (mut stream, peer) = match listener.accept().await {
                    Ok(s) => s,
                    Err(e) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: format!("Accept error on port {}: {}\n", config.port, e),
                            category: Some(types::OutputEventCategory::Stderr),
                            ..Default::default()
                        }));
                        continue;
                    }
                };

                {
                    let mut so = server_output.lock().unwrap();
                    let _ = so.send_event(Event::Output(events::OutputEventBody {
                        output: format!("Debug server: connection accepted from {}\n", peer),
                        category: Some(types::OutputEventCategory::Stdout),
                        ..Default::default()
                    }));
                }

                let request = match AspServer::read_request(&mut stream).await {
                    Ok(r) => r,
                    Err(e) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: format!("Debug server: read_request error: {}\n", e),
                            category: Some(types::OutputEventCategory::Stderr),
                            ..Default::default()
                        }));
                        continue;
                    }
                };

                {
                    let mut so = server_output.lock().unwrap();
                    let _ = so.send_event(Event::Output(events::OutputEventBody {
                        output: format!("Debug server: calling process_request (path={:?})\n", request.path),
                        category: Some(types::OutputEventCategory::Stdout),
                        ..Default::default()
                    }));
                }

                let response = match tokio::time::timeout(
                    std::time::Duration::from_secs(10),
                    AspServer::process_request(
                        request,
                        &server.handler_chain,
                        &folder,
                        &default_document,
                        &server.store,
                        Some(Arc::clone(&debugger)),
                        config.directory_listing,
                    ),
                )
                .await
                {
                    Ok(Ok(r)) => r,
                    Ok(Err(e)) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: format!("Debug server: process error: {}\n", e),
                            category: Some(types::OutputEventCategory::Stderr),
                            ..Default::default()
                        }));
                        asperger::asp::server::HttpResponse {
                            status_line: "500 Internal Server Error".to_string(),
                            content_type: "text/plain".to_string(),
                            body: format!("Error: {}", e).into_bytes(),
                            extra_headers: Vec::new(),
                        }
                    }
                    Err(_) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: "Debug server: process_request TIMEOUT (10s) — debugger likely blocked on check()\n".to_string(),
                            category: Some(types::OutputEventCategory::Stderr),
                            ..Default::default()
                        }));
                        asperger::asp::server::HttpResponse {
                            status_line: "504 Gateway Timeout".to_string(),
                            content_type: "text/plain".to_string(),
                            body: "Debug server timeout — debugger blocked indefinitely".to_string().into_bytes(),
                            extra_headers: Vec::new(),
                        }
                    }
                };

                match AspServer::write_response(&mut stream, &response).await {
                    Ok(()) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: format!(
                                "Debug server: wrote response ({} bytes, status={:?})\n",
                                response.body.len(),
                                response.status_line
                            ),
                            category: Some(types::OutputEventCategory::Stdout),
                            ..Default::default()
                        }));
                    }
                    Err(e) => {
                        let mut so = server_output.lock().unwrap();
                        let _ = so.send_event(Event::Output(events::OutputEventBody {
                            output: format!("Debug server: write_response error: {}\n", e),
                            category: Some(types::OutputEventCategory::Stderr),
                            ..Default::default()
                        }));
                    }
                }
            }
        });
    }));

    if let Err(panic_err) = result {
        let msg = if let Some(s) = panic_err.downcast_ref::<&str>() {
            s.to_string()
        } else if let Some(s) = panic_err.downcast_ref::<String>() {
            s.clone()
        } else {
            "unknown panic".to_string()
        };
        let mut so = server_output.lock().unwrap();
        let _ = so.send_event(Event::Output(events::OutputEventBody {
            output: format!("Debug server: PANIC in HTTP thread: {}\n", msg),
            category: Some(types::OutputEventCategory::Stderr),
            ..Default::default()
        }));
    }
}
