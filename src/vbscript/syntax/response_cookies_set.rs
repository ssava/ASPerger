use super::VBSyntax;
use crate::vbscript::asp_objects::to_cookie_string;
use crate::vbscript::execution_context::CookieEntry;
use crate::vbscript::expr::{evaluate, Expr};
use crate::vbscript::{vbs_error::VBSError, ExecutionContext};

#[derive(Clone)]
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
        let entry = context
            .response
            .cookies
            .entry(name.clone())
            .or_insert_with(|| CookieEntry {
                value: val.clone(),
                ..Default::default()
            });
        entry.value = val.clone();
        context.response.extra_headers.push((
            "Set-Cookie".to_string(),
            to_cookie_string(&name, entry),
        ));
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}

#[derive(Clone)]
pub struct ResponseCookiesSetProp {
    key: Expr,
    property: String,
    value: Expr,
}

impl ResponseCookiesSetProp {
    pub fn new(key: Expr, property: String, value: Expr) -> Self {
        ResponseCookiesSetProp {
            key,
            property,
            value,
        }
    }
}

impl VBSyntax for ResponseCookiesSetProp {
    fn execute(&self, context: &mut ExecutionContext) -> Result<(), VBSError> {
        let key = evaluate(&self.key, context)?;
        let value = evaluate(&self.value, context)?;
        let name = crate::vbscript::value_utils::to_arg_string(&key);
        let entry = context
            .response
            .cookies
            .entry(name.clone())
            .or_default();
        let str_val = crate::vbscript::value_utils::to_arg_string(&value);
        match self.property.to_uppercase().as_str() {
            "EXPIRES" => entry.expires = str_val,
            "DOMAIN" => entry.domain = str_val,
            "PATH" => entry.path = str_val,
            "SECURE" => entry.secure = crate::vbscript::value_utils::to_boolean(&value),
            _ => entry.value = str_val,
        }
        context.response.extra_headers.push((
            "Set-Cookie".to_string(),
            to_cookie_string(&name, entry),
        ));
        Ok(())
    }

    fn clone_box(&self) -> Box<dyn VBSyntax> {
        Box::new(self.clone())
    }
}
