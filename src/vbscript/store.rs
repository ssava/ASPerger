//! Thread-safe shared storage for session and application data,
//! replacing previous global statics with explicit dependency injection.

use std::sync::{Arc, Mutex, MutexGuard};

use ahash::AHashMap;

use super::value::VBValue;

/// Shared store containing session data, application data, and an application-level mutex lock.
pub struct Store {
    sessions: Mutex<AHashMap<String, AHashMap<String, VBValue>>>,
    apps: Mutex<AHashMap<String, VBValue>>,
    app_lock: Mutex<()>,
}

impl Store {
    /// Create a new shared store wrapped in `Arc`.
    pub fn new() -> Arc<Self> {
        Arc::new(Store {
            sessions: Mutex::new(AHashMap::new()),
            apps: Mutex::new(AHashMap::new()),
            app_lock: Mutex::new(()),
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
    pub fn lock_app(&self) -> MutexGuard<'_, ()> {
        self.app_lock.lock().unwrap()
    }

    pub fn clear_apps(&self) {
        self.apps.lock().unwrap().clear();
    }
}
