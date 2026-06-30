//! VBScript block / control-flow parsing and execution.

mod block_types;
mod block_parse;
pub(crate) mod block_exec;

pub use block_types::{BlockStatement, CaseClause, ElseIfBlock, UserDefinedFunction};
pub use block_parse::parse_blocks;
pub(crate) use block_exec::execute_user_defined_function;

