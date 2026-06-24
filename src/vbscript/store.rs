//! Thread-safe shared storage for session and application data,
//! Global.asa state, and application-scoped static objects.

use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU64, Ordering};
use std::sync::{Arc, Condvar, Mutex, MutexGuard};

use ahash::AHashMap;

use super::value::VBValue;

use crate::asp::global_asa::GlobalAsa;

/// State tracked for the Application.Lock/Unlock mechanism.
struct AppLockInfo {
    locked: bool,
    owner_id: u64,
}

/// Shared store containing session data, application data, an application-level
/// mutex lock, Global.asa event handlers, and application-scoped static objects.
pub struct Store {
    sessions: Mutex<AHashMap<String, AHashMap<String, VBValue>>>,
    apps: Mutex<AHashMap<String, VBValue>>,
    app_lock_mtx: Mutex<AppLockInfo>,
    app_lock_cv: Condvar,
    /// Counter for generating unique per-request IDs.
    next_request_id: AtomicU64,
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
            app_lock_mtx: Mutex::new(AppLockInfo { locked: false, owner_id: 0 }),
            app_lock_cv: Condvar::new(),
            next_request_id: AtomicU64::new(1),
            global_asa: Mutex::new(None),
            app_started: AtomicBool::new(false),
            app_static_objects: Mutex::new(AHashMap::new()),
            session_timeout_minutes: AtomicI32::new(20),
        })
    }

    /// Allocate a unique request ID for this request.
    pub fn allocate_request_id(&self) -> u64 {
        self.next_request_id.fetch_add(1, Ordering::Relaxed)
    }

    /// Lock and return a mutable guard to the session store.
    pub fn lock_sessions(&self) -> MutexGuard<'_, AHashMap<String, AHashMap<String, VBValue>>> {
        self.sessions.lock().unwrap()
    }

    /// Lock and return a mutable guard to the application store.
    pub fn lock_apps(&self) -> MutexGuard<'_, AHashMap<String, VBValue>> {
        self.apps.lock().unwrap()
    }

    /// Block until the application lock is released by another request, then return.
    /// If the lock is held by the current owner_id, returns immediately.
    /// Must be called before every Application variable access (indexed_get/set, CONTENTS methods).
    pub fn wait_for_app_unlock(&self, my_owner_id: u64) {
        let mut info = self.app_lock_mtx.lock().unwrap();
        while info.locked && info.owner_id != my_owner_id {
            info = self.app_lock_cv.wait(info).unwrap();
        }
    }

    /// Acquire the application lock (blocks until available).
    /// Returns true if the lock was acquired, false if already held by this owner.
    pub fn lock_app_blocking(&self, my_owner_id: u64) -> bool {
        let mut info = self.app_lock_mtx.lock().unwrap();
        while info.locked {
            if info.owner_id == my_owner_id {
                return false; // already locked by us
            }
            info = self.app_lock_cv.wait(info).unwrap();
        }
        info.locked = true;
        info.owner_id = my_owner_id;
        true
    }

    /// Release the application lock if held by this owner.
    /// Returns true if the lock was released, false if not held.
    pub fn unlock_app(&self, my_owner_id: u64) -> bool {
        let mut info = self.app_lock_mtx.lock().unwrap();
        if info.locked && info.owner_id == my_owner_id {
            info.locked = false;
            self.app_lock_cv.notify_all();
            return true;
        }
        false
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_store_session_write_read() {
        let store = Store::new();
        let key = "test_session".to_string();
        {
            let mut sessions = store.lock_sessions();
            let mut data = AHashMap::new();
            data.insert("foo".to_string(), VBValue::String("bar".to_string()));
            sessions.insert(key.clone(), data);
        }
        {
            let sessions = store.lock_sessions();
            let data = sessions.get(&key).unwrap();
            assert_eq!(data.get("foo").unwrap().to_string(), "bar");
        }
    }

    #[test]
    fn test_store_session_remove() {
        let store = Store::new();
        let sid = "SES1";
        {
            let mut sessions = store.lock_sessions();
            sessions.insert(sid.to_string(), AHashMap::new());
        }
        assert_eq!(store.session_count(), 1);
        store.remove_session(sid);
        assert_eq!(store.session_count(), 0);
    }

    #[test]
    fn test_store_app_write_read() {
        let store = Store::new();
        let owner = store.allocate_request_id();
        {
            let mut apps = store.lock_apps();
            apps.insert("counter".to_string(), VBValue::Number(42.0));
        }
        store.wait_for_app_unlock(owner);
        let apps = store.lock_apps();
        let val = apps.get("counter").unwrap();
        match val {
            VBValue::Number(n) => assert!((n - 42.0).abs() < 1e-10),
            _ => panic!("expected Number"),
        }
    }

    #[test]
    fn test_store_app_lock_unlock_blocks_other() {
        let store = Store::new();
        let id_a = store.allocate_request_id();
        let id_b = store.allocate_request_id();

        // A acquires the lock
        assert!(store.lock_app_blocking(id_a));

        // B tries to acquire — would block, so use try_lock semantics via a thread
        // Instead, verify B must wait by checking lock state
        let info = store.app_lock_mtx.lock().unwrap();
        assert!(info.locked);
        assert_eq!(info.owner_id, id_a);
        drop(info);

        // A releases
        assert!(store.unlock_app(id_a));

        // B can now acquire
        assert!(store.lock_app_blocking(id_b));
        assert!(store.unlock_app(id_b));
    }

    #[test]
    fn test_store_app_lock_reentrant() {
        let store = Store::new();
        let id = store.allocate_request_id();
        assert!(store.lock_app_blocking(id));
        // Reentrant lock returns false (already held)
        assert!(!store.lock_app_blocking(id));
        assert!(store.unlock_app(id));
    }

    #[test]
    fn test_store_multiple_sessions_independent() {
        let store = Store::new();
        {
            let mut sessions = store.lock_sessions();
            let mut a = AHashMap::new();
            a.insert("key".to_string(), VBValue::String("val_a".to_string()));
            sessions.insert("A".to_string(), a);
            let mut b = AHashMap::new();
            b.insert("key".to_string(), VBValue::String("val_b".to_string()));
            sessions.insert("B".to_string(), b);
        }
        {
            let sessions = store.lock_sessions();
            assert_eq!(
                sessions.get("A").unwrap().get("key").unwrap().to_string(),
                "val_a"
            );
            assert_eq!(
                sessions.get("B").unwrap().get("key").unwrap().to_string(),
                "val_b"
            );
        }
    }
}
