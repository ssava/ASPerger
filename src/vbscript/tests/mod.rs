use crate::asp::parser::AspParser;
use crate::vbscript::expr::{evaluate, parse_expression, BinOp, Expr, UnaryOp};
use crate::vbscript::syntax::{Assignment, Dim, ResponseWrite, VBSyntax};
use crate::vbscript::{ExecutionContext, TokenType, Tokenizer, VBScriptInterpreter, VBValue};
use chrono::{Datelike, Timelike};
use std::fs;
use std::path::Path;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;

pub(crate) fn tmp_path(name: &str) -> String {
    let p = std::env::temp_dir().join(format!("asperger_test_{}", name));
    p.to_str().unwrap().to_string()
}

pub(crate) fn cleanup_path(path: &str) {
    let p = Path::new(path);
    if p.is_file() {
        let _ = std::fs::remove_file(p);
    } else if p.is_dir() {
        let _ = std::fs::remove_dir_all(p);
    }
}

static TEST_COUNTER: AtomicU64 = AtomicU64::new(1);

pub(crate) fn tmp_asp_dir() -> std::path::PathBuf {
    let counter = TEST_COUNTER.fetch_add(1, Ordering::Relaxed);
    let dir = std::env::temp_dir().join(format!("asperger_test_{}_{}", std::process::id(), counter));
    let _ = std::fs::create_dir_all(&dir);
    dir
}

pub(crate) fn write_asp(dir: &std::path::Path, name: &str, content: &str) {
    std::fs::write(dir.join(name), format!("<%@ LANGUAGE=VBScript %>{}", content)).unwrap();
}

pub(crate) fn cleanup_dir(dir: &std::path::Path) {
    if dir.exists() {
        for entry in std::fs::read_dir(dir).unwrap() {
            if let Ok(e) = entry {
                let _ = std::fs::remove_file(e.path());
            }
        }
        let _ = std::fs::remove_dir(dir);
    }
}

pub(crate) mod tokenizer;
pub(crate) mod expressions;
pub(crate) mod blocks;
pub(crate) mod builtins;
pub(crate) mod objects;
pub(crate) mod fso;
pub(crate) mod asp;
