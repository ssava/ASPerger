use super::VBSyntax;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

pub struct ResponseCookiesSet {
    key: Expr,
    value: Expr,
}

impl ResponseCookiesSet {
    pub fn new(key: Expr, value: Expr) -> Self {
        ResponseCookiesSet { key, value }
    }
}

impl VBSyntax for ResponseCookiesSet {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let key = evaluate(&self.key, context)?;
        let value = evaluate(&self.value, context)?;
        let name = crate::vbscript::value_utils::to_arg_string(&key);
        let val = crate::vbscript::value_utils::to_arg_string(&value);
        context.response.cookies.insert(name.clone(), val.clone());
        context.response.extra_headers.push((
            "Set-Cookie".to_string(),
            format!("{}={}; path=/", name, val),
        ));
        Ok(())
    }
}
