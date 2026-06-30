//! VBScript block / control-flow parsing and compilation.

mod block_types;
mod block_parse;

pub use block_types::{BlockStatement, CaseClause, ElseIfBlock, UserDefinedFunction};
pub(crate) use block_parse::first_non_ws;
pub use block_parse::parse_blocks;

