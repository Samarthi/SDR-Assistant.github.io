// Lightweight shim to run the extension code in a plain web page.
// Implements the subset of chrome.* APIs used by background-web.js and popup-web.js.
(function initChromeShim() {
  if (window.chrome && window.chrome.runtime && window.chrome.storage) {
    // Already running in an environment that provides chrome.*; don't override.
    return;
  }

  const STORAGE_KEY = "__chrome_shim_storage__";
  let store = {};
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    store = raw ? JSON.parse(raw) : {};
  } catch (err) {
    console.warn("Failed to parse stored shim data; starting fresh", err);
    store = {};
  }
  const persistStore = () => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
    } catch (err) {
      console.warn("Failed to persist shim storage", err);
    }
  };

  const makeEvent = () => {
    const listeners = [];
    return {
      addListener(fn) {
        if (typeof fn === "function") listeners.push(fn);
      },
      removeListener(fn) {
        const idx = listeners.indexOf(fn);
        if (idx >= 0) listeners.splice(idx, 1);
      },
      hasListeners() {
        return listeners.length > 0;
      },
      _listeners: listeners,
    };
  };

  const alarmsEvent = makeEvent();
  const runtimeMessageEvent = makeEvent();
  const runtimeInstalledEvent = makeEvent();
  const runtimeStartupEvent = makeEvent();

  const alarms = {
    _timers: new Map(),
    create(name, opts = {}) {
      if (!name) return;
      const delayMs = Math.max(0, (opts.delayInMinutes || 0) * 60 * 1000);
      const periodMs = opts.periodInMinutes ? Math.max(0, opts.periodInMinutes * 60 * 1000) : null;

      const fire = () => {
        alarmsEvent._listeners.forEach((fn) => {
          try {
            fn({ name });
          } catch (err) {
            console.warn("Alarm listener failed", err);
          }
        });
      };

      if (this._timers.has(name)) {
        clearTimeout(this._timers.get(name)?.timeoutId);
        clearInterval(this._timers.get(name)?.intervalId);
      }

      const timeoutId = setTimeout(() => {
        fire();
        if (periodMs) {
          const intervalId = setInterval(fire, periodMs);
          this._timers.set(name, { intervalId });
        }
      }, delayMs || 0);

      this._timers.set(name, { timeoutId });
    },
    clear(name) {
      const entry = this._timers.get(name);
      if (!entry) return;
      clearTimeout(entry.timeoutId);
      clearInterval(entry.intervalId);
      this._timers.delete(name);
    },
    onAlarm: alarmsEvent,
  };

  const storage = {
    local: {
      async get(keys) {
        if (keys === null || keys === undefined) {
          return { ...store };
        }
        if (typeof keys === "string") {
          return { [keys]: store[keys] };
        }
        if (Array.isArray(keys)) {
          const result = {};
          keys.forEach((k) => {
            result[k] = store[k];
          });
          return result;
        }
        if (typeof keys === "object") {
          const result = {};
          Object.keys(keys).forEach((k) => {
            result[k] = store[k] !== undefined ? store[k] : keys[k];
          });
          return result;
        }
        return { ...store };
      },
      async set(obj = {}) {
        Object.assign(store, obj);
        persistStore();
      },
      async remove(keys) {
        const list = Array.isArray(keys) ? keys : [keys];
        list.forEach((k) => {
          delete store[k];
        });
        persistStore();
      },
      async clear() {
        store = {};
        persistStore();
      },
    },
  };

  const runtime = {
    lastError: null,
    onMessage: runtimeMessageEvent,
    onInstalled: runtimeInstalledEvent,
    onStartup: runtimeStartupEvent,
    getURL: (path) => path,
    sendMessage(message, callback) {
      const listeners = runtimeMessageEvent._listeners;
      if (!listeners.length) {
        runtime.lastError = new Error("No message listeners registered");
        if (callback) callback(undefined);
        return Promise.resolve(undefined);
      }

      runtime.lastError = null;
      return new Promise((resolve) => {
        let responded = false;
        const sendResponse = (resp) => {
          if (responded) return;
          responded = true;
          if (callback) {
            try {
              callback(resp);
            } catch (err) {
              console.warn("Callback threw after sendResponse", err);
            }
          }
          resolve(resp);
        };

        listeners.forEach((fn) => {
          try {
            const result = fn(message, { id: "shim" }, sendResponse);
            if (result && typeof result.then === "function") {
              result.catch((err) => {
                console.warn("Async listener error", err);
              });
            }
          } catch (err) {
            console.warn("Message listener error", err);
            runtime.lastError = err;
          }
        });
      });
    },
  };

  const action = {
    onClicked: makeEvent(),
  };

  const tabs = {
    create({ url }) {
      if (!url) return;
      try {
        window.open(url, "_blank");
      } catch (err) {
        console.warn("tabs.create failed", err);
      }
    },
  };

  const shim = {
    alarms,
    storage,
    runtime,
    action,
    tabs,
    // Placeholder namespaces for parity; unused in the web build.
    downloads: {},
    scripting: {},
  };

  window.chrome = Object.assign(window.chrome || {}, shim);

  // Trigger installed/startup listeners once to mirror extension bootstrapping.
  setTimeout(() => {
    runtimeInstalledEvent._listeners.forEach((fn) => {
      try {
        fn({ reason: "install" });
      } catch (err) {
        console.warn("onInstalled listener failed", err);
      }
    });
    runtimeStartupEvent._listeners.forEach((fn) => {
      try {
        fn();
      } catch (err) {
        console.warn("onStartup listener failed", err);
      }
    });
  }, 0);
})();
