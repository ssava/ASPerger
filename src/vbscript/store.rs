//! Thread-safe shared storage for session and application data,
//! Global.asa state, and application-scoped static objects.

use std::sync::atomic::{AtomicBool, AtomicI32, Ordering};
use std::sync::{Arc, Mutex, MutexGuard};

use ahash::AHashMap;

use super::value::VBValue;

use crate::asp::global_asa::GlobalAsa;

/// Shared store containing session data, application data, an application-level
/// mutex lock, Global.asa event handlers, and application-scoped static objects.
pub struct Store {
    sessions: Mutex<AHashMap<String, AHashMap<String, VBValue>>>,
    apps: Mutex<AHashMap<String, VBValue>>,
    app_lock_state: Mutex<bool>,
    /// Parsed Global.asa data, loaded once at startup.
    pub global_asa: Mutex<Option<GlobalAsa>>,
    /// Whether Application_OnStart has been triggered.
    pub app_started: AtomicBool,
    /// Application-scoped objects from <OBJECT SCOPE="Application"> declarations.
    pub app_static_objects: Mutex<AHashMap<String, VBValue>>,
    /// Session timeout in minutes (default 20).
    pub session_timeout_minutes: AtomicI32,
}

impl Store {
    /// Create a new shared store wrapped in `Arc`.
    pub fn new() -> Arc<Self> {
        Arc::new(Store {
            sessions: Mutex::new(AHashMap::new()),
            apps: Mutex::new(AHashMap::new()),
            app_lock_state: Mutex::new(false),
            global_asa: Mutex::new(None),
            app_started: AtomicBool::new(false),
            app_static_objects: Mutex::new(AHashMap::new()),
            session_timeout_minutes: AtomicI32::new(20),
        })
    }

    /// Lock and return a mutable guard to the session store.
    pub fn lock_sessions(&self) -> MutexGuard<'_, AHashMap<String, AHashMap<String, VBValue>>> {
        self.sessions.lock().unwrap()
    }

    /// Lock and return a mutable guard to the application store.
    pub fn lock_apps(&self) -> MutexGuard<'_, AHashMap<String, VBValue>> {
        self.apps.lock().unwrap()
    }

    /// Lock the application mutex (used by Application.Lock / Unlock).
    pub fn lock_app(&self) -> MutexGuard<'_, bool> {
        self.app_lock_state.lock().unwrap()
    }

    pub fn clear_apps(&self) {
        self.apps.lock().unwrap().clear();
    }

    /// Check if Application_OnStart has been fired and mark it as fired.
    pub fn try_start_application(&self) -> bool {
        self.app_started
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .is_ok()
    }

    /// Load Global.asa into the store.
    pub fn set_global_asa(&self, global_asa: GlobalAsa) {
        *self.global_asa.lock().unwrap() = Some(global_asa);
    }

    /// Get a reference to the parsed Global.asa if available.
    pub fn get_global_asa(&self) -> Option<GlobalAsa> {
        self.global_asa.lock().unwrap().clone()
    }

    /// Get the number of active sessions.
    pub fn session_count(&self) -> usize {
        self.sessions.lock().unwrap().len()
    }

    /// Remove a session (used by Abandon and timeout sweep).
    pub fn remove_session(&self, session_id: &str) {
        self.sessions.lock().unwrap().remove(&session_id.to_uppercase());
    }

    /// Update the last-access timestamp for a session.
    pub fn touch_session(&self, session_id: &str) {
        let mut sessions = self.sessions.lock().unwrap();
        if let Some(data) = sessions.get_mut(&session_id.to_uppercase()) {
            data.insert(
                "__LAST_ACCESS__".to_string(),
                VBValue::String(chrono::Utc::now().to_rfc3339()),
            );
        }
    }
}
