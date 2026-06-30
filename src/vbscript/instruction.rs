use std::fmt;

type ConstantIdx = u32;
type LocalSlot = usize;
type CodeOffset = i32;

#[derive(Debug, Clone, PartialEq)]
pub enum Instruction {
    // -- Constants --
    LoadConst(ConstantIdx),
    LoadNil,
    LoadTrue,
    LoadFalse,
    LoadEmpty,

    // -- Variables --
    LoadLocal(LocalSlot),
    StoreLocal(LocalSlot),
    LoadGlobal(ConstantIdx),
    StoreGlobal(ConstantIdx),

    // -- Unary --
    Neg,
    Not,

    // -- Binary arithmetic --
    Add,
    Sub,
    Mul,
    Div,
    IntDiv,
    Mod,
    Pow,

    // -- String --
    Concat,

    // -- Comparison --
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Is,
    Like,

    // -- Logical --
    And,
    Or,
    Xor,
    Imp,
    Eqv,

    // -- Objects --
    GetProp(ConstantIdx),
    SetProp(ConstantIdx),
    SetPropLocal(LocalSlot, ConstantIdx),
    SetPropGlobal(ConstantIdx, ConstantIdx),
    CallMethod(ConstantIdx, u8),
    CallMethodLocal(LocalSlot, ConstantIdx, u8),
    CallMethodGlobal(ConstantIdx, ConstantIdx, u8),
    IndexGet,
    IndexSet,
    NewObject(ConstantIdx),

    // -- Arrays --
    NewArray(u8),
    ReDim(LocalSlot, u8, bool),
    IndexStoreLocal(LocalSlot),
    IndexStoreGlobal(ConstantIdx),
    IndexStoreLocalMulti(LocalSlot, u8),
    IndexStoreGlobalMulti(ConstantIdx, u8),

    // -- Functions --
    Call(ConstantIdx, u8),
    CallLocal(LocalSlot, u8),
    Return(u8),

    // -- Control flow --
    Jump(CodeOffset),
    JumpIfFalse(CodeOffset),
    JumpIfTrue(CodeOffset),

    // -- Loops --
    ForPrep(LocalSlot, CodeOffset),
    ForStep(LocalSlot, CodeOffset),
    ForEachPrep(LocalSlot, CodeOffset),
    ForEachStep(LocalSlot, CodeOffset),

    // -- Exit signals --
    ExitFor,
    ExitDo,
    ExitFunction,
    ExitSub,

    // -- Scope --
    SelectStore,
    LoadSelectValue,
    SelectCompare,
    SelectClear,
    WithStart,
    WithEnd,

    // -- Error handling --
    OnErrorResumeNext,
    OnErrorGoto0,
    Raise(ConstantIdx),

    // -- Debug --
    DebugLine(u32),

    // -- ASP-specific --
    ResponseWrite,
    ResponseEnd,
    ServerExecute(ConstantIdx),
    ServerTransfer(ConstantIdx),

    // -- Variable management --
    Erase(LocalSlot),
    EraseGlobal(ConstantIdx),
}

impl fmt::Display for Instruction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Instruction::LoadConst(i) => write!(f, "LoadConst {}", i),
            Instruction::LoadNil => write!(f, "LoadNil"),
            Instruction::LoadTrue => write!(f, "LoadTrue"),
            Instruction::LoadFalse => write!(f, "LoadFalse"),
            Instruction::LoadEmpty => write!(f, "LoadEmpty"),
            Instruction::LoadLocal(s) => write!(f, "LoadLocal {}", s),
            Instruction::StoreLocal(s) => write!(f, "StoreLocal {}", s),
            Instruction::LoadGlobal(i) => write!(f, "LoadGlobal {}", i),
            Instruction::StoreGlobal(i) => write!(f, "StoreGlobal {}", i),
            Instruction::Neg => write!(f, "Neg"),
            Instruction::Not => write!(f, "Not"),
            Instruction::Add => write!(f, "Add"),
            Instruction::Sub => write!(f, "Sub"),
            Instruction::Mul => write!(f, "Mul"),
            Instruction::Div => write!(f, "Div"),
            Instruction::IntDiv => write!(f, "IntDiv"),
            Instruction::Mod => write!(f, "Mod"),
            Instruction::Pow => write!(f, "Pow"),
            Instruction::Concat => write!(f, "Concat"),
            Instruction::Eq => write!(f, "Eq"),
            Instruction::Ne => write!(f, "Ne"),
            Instruction::Lt => write!(f, "Lt"),
            Instruction::Le => write!(f, "Le"),
            Instruction::Gt => write!(f, "Gt"),
            Instruction::Ge => write!(f, "Ge"),
            Instruction::Is => write!(f, "Is"),
            Instruction::Like => write!(f, "Like"),
            Instruction::And => write!(f, "And"),
            Instruction::Or => write!(f, "Or"),
            Instruction::Xor => write!(f, "Xor"),
            Instruction::Imp => write!(f, "Imp"),
            Instruction::Eqv => write!(f, "Eqv"),
            Instruction::GetProp(i) => write!(f, "GetProp {}", i),
            Instruction::SetProp(i) => write!(f, "SetProp {}", i),
            Instruction::SetPropLocal(s, i) => write!(f, "SetPropLocal {} {}", s, i),
            Instruction::SetPropGlobal(o, i) => write!(f, "SetPropGlobal {} {}", o, i),
            Instruction::CallMethod(i, n) => write!(f, "CallMethod {} {}", i, n),
            Instruction::CallMethodLocal(s, i, n) => write!(f, "CallMethodLocal {} {} {}", s, i, n),
            Instruction::CallMethodGlobal(g, i, n) => write!(f, "CallMethodGlobal {} {} {}", g, i, n),
            Instruction::IndexGet => write!(f, "IndexGet"),
            Instruction::IndexSet => write!(f, "IndexSet"),
            Instruction::NewObject(i) => write!(f, "NewObject {}", i),
            Instruction::NewArray(n) => write!(f, "NewArray {}", n),
            Instruction::IndexStoreLocal(s) => write!(f, "IndexStoreLocal {}", s),
            Instruction::IndexStoreGlobal(i) => write!(f, "IndexStoreGlobal {}", i),
            Instruction::IndexStoreLocalMulti(s, n) => write!(f, "IndexStoreLocalMulti {} {}", s, n),
            Instruction::IndexStoreGlobalMulti(i, n) => write!(f, "IndexStoreGlobalMulti {} {}", i, n),
            Instruction::ReDim(s, n, p) => write!(f, "ReDim {} {} {}", s, n, p),
            Instruction::Call(i, n) => write!(f, "Call {} {}", i, n),
            Instruction::CallLocal(s, n) => write!(f, "CallLocal {} {}", s, n),
            Instruction::Return(n) => write!(f, "Return {}", n),
            Instruction::Jump(o) => write!(f, "Jump {}", o),
            Instruction::JumpIfFalse(o) => write!(f, "JumpIfFalse {}", o),
            Instruction::JumpIfTrue(o) => write!(f, "JumpIfTrue {}", o),
            Instruction::ForPrep(s, o) => write!(f, "ForPrep {} {}", s, o),
            Instruction::ForStep(s, o) => write!(f, "ForStep {} {}", s, o),
            Instruction::ForEachPrep(s, o) => write!(f, "ForEachPrep {} {}", s, o),
            Instruction::ForEachStep(s, o) => write!(f, "ForEachStep {} {}", s, o),
            Instruction::ExitFor => write!(f, "ExitFor"),
            Instruction::ExitDo => write!(f, "ExitDo"),
            Instruction::ExitFunction => write!(f, "ExitFunction"),
            Instruction::ExitSub => write!(f, "ExitSub"),
            Instruction::SelectStore => write!(f, "SelectStore"),
            Instruction::LoadSelectValue => write!(f, "LoadSelectValue"),
            Instruction::SelectCompare => write!(f, "SelectCompare"),
            Instruction::SelectClear => write!(f, "SelectClear"),
            Instruction::WithStart => write!(f, "WithStart"),
            Instruction::WithEnd => write!(f, "WithEnd"),
            Instruction::OnErrorResumeNext => write!(f, "OnErrorResumeNext"),
            Instruction::OnErrorGoto0 => write!(f, "OnErrorGoto0"),
            Instruction::Raise(i) => write!(f, "Raise {}", i),
            Instruction::DebugLine(l) => write!(f, "DebugLine {}", l),
            Instruction::ResponseWrite => write!(f, "ResponseWrite"),
            Instruction::ResponseEnd => write!(f, "ResponseEnd"),
            Instruction::ServerExecute(i) => write!(f, "ServerExecute {}", i),
            Instruction::ServerTransfer(i) => write!(f, "ServerTransfer {}", i),
            Instruction::Erase(s) => write!(f, "Erase {}", s),
            Instruction::EraseGlobal(i) => write!(f, "EraseGlobal {}", i),
        }
    }
}
