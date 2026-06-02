use std::sync::{Arc, Mutex, MutexGuard};

use ahash::AHashMap;

use super::value::VBValue;

pub struct Store {
    sessions: Mutex<AHashMap<String, AHashMap<String, VBValue>>>,
    apps: Mutex<AHashMap<String, VBValue>>,
    app_lock: Mutex<()>,
}

impl Store {
    pub fn new() -> Arc<Self> {
        Arc::new(Store {
            sessions: Mutex::new(AHashMap::new()),
            apps: Mutex::new(AHashMap::new()),
            app_lock: Mutex::new(()),
        })
    }

    pub fn lock_sessions(
        &self,
    ) -> MutexGuard<'_, AHashMap<String, AHashMap<String, VBValue>>> {
        self.sessions.lock().unwrap()
    }

    pub fn lock_apps(&self) -> MutexGuard<'_, AHashMap<String, VBValue>> {
        self.apps.lock().unwrap()
    }

    pub fn lock_app(&self) -> MutexGuard<'_, ()> {
        self.app_lock.lock().unwrap()
    }

    pub fn clear_apps(&self) {
        self.apps.lock().unwrap().clear();
    }
}
