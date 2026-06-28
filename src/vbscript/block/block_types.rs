use crate::vbscript::expr::Expr;
use crate::vbscript::syntax::VBSyntax;
use crate::vbscript::vbs_error::VBSError;
use crate::vbscript::Token;

/// Parsed VBScript statement, produced by `parse_blocks`.
///
/// Each variant corresponds to a VBScript control-flow or declaration construct.
/// The `line` field (present on all compound variants) stores the source line
/// number used for debugger breakpoint matching.
pub enum BlockStatement {
    Syntax(Box<dyn VBSyntax>, usize),
    Unrecognized(VBSError, String, usize),
    If {
        line: usize,
        condition: Expr,
        then_body: Vec<BlockStatement>,
        else_if_blocks: Vec<ElseIfBlock>,
        else_body: Option<Vec<BlockStatement>>,
    },
    For {
        line: usize,
        counter: String,
        start: Expr,
        end: Expr,
        step: Option<Expr>,
        body: Vec<BlockStatement>,
    },
    While {
        line: usize,
        condition: Expr,
        body: Vec<BlockStatement>,
    },
    Do {
        line: usize,
        body: Vec<BlockStatement>,
        condition: Option<Expr>,
        is_until: bool,
        is_post_test: bool,
    },
    ForEach {
        line: usize,
        element: String,
        group: Expr,
        body: Vec<BlockStatement>,
    },
    FunctionDef {
        line: usize,
        name: String,
        params: Vec<String>,
        body_lines: Vec<Vec<Token>>,
    },
    SubDef {
        line: usize,
        name: String,
        params: Vec<String>,
        body_lines: Vec<Vec<Token>>,
    },
    SelectCase {
        line: usize,
        expression: Expr,
        cases: Vec<CaseClause>,
        else_body: Option<Vec<BlockStatement>>,
    },
    ClassDef {
        line: usize,
        name: String,
        body_lines: Vec<Vec<Token>>,
    },
    With {
        line: usize,
        object: Expr,
        body: Vec<BlockStatement>,
    },
    ExitFor(usize),
    ExitDo(usize),
    ExitFunction(usize),
    ExitSub(usize),
}

impl Clone for BlockStatement {
    fn clone(&self) -> Self {
        match self {
            BlockStatement::Syntax(s, line) => BlockStatement::Syntax(s.clone_box(), *line),
            BlockStatement::Unrecognized(err, s, line) => {
                BlockStatement::Unrecognized(err.clone(), s.clone(), *line)
            }
            BlockStatement::If {
                line,
                condition,
                then_body,
                else_if_blocks,
                else_body,
            } => BlockStatement::If {
                line: *line,
                condition: condition.clone(),
                then_body: then_body.clone(),
                else_if_blocks: else_if_blocks.clone(),
                else_body: else_body.clone(),
            },
            BlockStatement::For {
                line,
                counter,
                start,
                end,
                step,
                body,
            } => BlockStatement::For {
                line: *line,
                counter: counter.clone(),
                start: start.clone(),
                end: end.clone(),
                step: step.clone(),
                body: body.clone(),
            },
            BlockStatement::While {
                line,
                condition,
                body,
            } => BlockStatement::While {
                line: *line,
                condition: condition.clone(),
                body: body.clone(),
            },
            BlockStatement::Do {
                line,
                body,
                condition,
                is_until,
                is_post_test,
            } => BlockStatement::Do {
                line: *line,
                body: body.clone(),
                condition: condition.clone(),
                is_until: *is_until,
                is_post_test: *is_post_test,
            },
            BlockStatement::ForEach {
                line,
                element,
                group,
                body,
            } => BlockStatement::ForEach {
                line: *line,
                element: element.clone(),
                group: group.clone(),
                body: body.clone(),
            },
            BlockStatement::FunctionDef {
                line,
                name,
                params,
                body_lines,
            } => BlockStatement::FunctionDef {
                line: *line,
                name: name.clone(),
                params: params.clone(),
                body_lines: body_lines.clone(),
            },
            BlockStatement::SubDef {
                line,
                name,
                params,
                body_lines,
            } => BlockStatement::SubDef {
                line: *line,
                name: name.clone(),
                params: params.clone(),
                body_lines: body_lines.clone(),
            },
            BlockStatement::SelectCase {
                line,
                expression,
                cases,
                else_body,
            } => BlockStatement::SelectCase {
                line: *line,
                expression: expression.clone(),
                cases: cases.clone(),
                else_body: else_body.clone(),
            },
            BlockStatement::ClassDef {
                line,
                name,
                body_lines,
            } => BlockStatement::ClassDef {
                line: *line,
                name: name.clone(),
                body_lines: body_lines.clone(),
            },
            BlockStatement::With {
                line,
                object,
                body,
            } => BlockStatement::With {
                line: *line,
                object: object.clone(),
                body: body.clone(),
            },
            BlockStatement::ExitFor(l) => BlockStatement::ExitFor(*l),
            BlockStatement::ExitDo(l) => BlockStatement::ExitDo(*l),
            BlockStatement::ExitFunction(l) => BlockStatement::ExitFunction(*l),
            BlockStatement::ExitSub(l) => BlockStatement::ExitSub(*l),
        }
    }
}

/// A single `ElseIf condition Then` clause inside an `If` block.
#[derive(Clone)]
pub struct ElseIfBlock {
    pub condition: Expr,
    pub body: Vec<BlockStatement>,
}

/// A single `Case values` clause inside a `Select Case` block.
/// `Case Is operator value` is encoded as `Expr::CaseComparison`.
#[derive(Clone)]
pub struct CaseClause {
    pub values: Vec<Expr>,
    pub body: Vec<BlockStatement>,
}

impl BlockStatement {
    pub fn line(&self) -> usize {
        match self {
            BlockStatement::Syntax(_, l) => *l,
            BlockStatement::Unrecognized(_, _, l) => *l,
            BlockStatement::If { line: l, .. } => *l,
            BlockStatement::For { line: l, .. } => *l,
            BlockStatement::While { line: l, .. } => *l,
            BlockStatement::Do { line: l, .. } => *l,
            BlockStatement::ForEach { line: l, .. } => *l,
            BlockStatement::FunctionDef { line: l, .. } => *l,
            BlockStatement::SubDef { line: l, .. } => *l,
            BlockStatement::SelectCase { line: l, .. } => *l,
            BlockStatement::ClassDef { line: l, .. } => *l,
            BlockStatement::With { line: l, .. } => *l,
            BlockStatement::ExitFor(l) => *l,
            BlockStatement::ExitDo(l) => *l,
            BlockStatement::ExitFunction(l) => *l,
            BlockStatement::ExitSub(l) => *l,
        }
    }
}

/// A user-defined `Sub` or `Function` parsed from source.
///
/// Function bodies are stored as raw token lines so they can be re-parsed
/// into `BlockStatement`s on each call (VBScript allows redefinition).
/// The cached parsed bodies are stored separately in `ExecutionContext::function_bodies`.
#[derive(Clone)]
pub struct UserDefinedFunction {
    pub name: String,
    pub params: Vec<String>,
    pub body_lines: Vec<Vec<Token>>,
    pub is_function: bool,
}
