//! ASPerger — a lightweight classic ASP (VBScript) server written in Rust.
//!
//! This crate provides an HTTP server capable of parsing and executing
//! ASP pages using a built-in VBScript interpreter, along with support
//! for ASP intrinsic objects (Request, Response, Session, Server, Application),
//! COM objects (ADODB, Scripting.FileSystemObject, etc.), and a DAP-based
//! debugger for VS Code integration.

pub mod asp;
pub mod vbscript;
