(function exportTransformWorkerScope() {
  const blocked = () => {
    throw new Error("Blocked API in transform worker.");
  };

  const safePostMessage = self.postMessage.bind(self);
  self.fetch = blocked;
  self.XMLHttpRequest = blocked;
  self.WebSocket = blocked;
  self.importScripts = blocked;
  self.postMessage = blocked;
  self.close = blocked;

  function toString(value) {
    if (value === undefined || value === null) return "";
    if (typeof value === "string") return value;
    if (typeof value === "number" || typeof value === "boolean") return String(value);
    try {
      return JSON.stringify(value);
    } catch (err) {
      return String(value);
    }
  }

  function coalesce() {
    for (let i = 0; i < arguments.length; i += 1) {
      const value = arguments[i];
      if (value !== undefined && value !== null && String(value).length) {
        return value;
      }
    }
    return "";
  }

  function join(value, sep) {
    const delimiter = typeof sep === "string" ? sep : ", ";
    if (Array.isArray(value)) {
      return value.map((item) => toString(item)).filter(Boolean).join(delimiter);
    }
    return toString(value);
  }

  function truncate(value, len) {
    const text = toString(value);
    const limit = Math.max(0, Number(len) || 0);
    if (!limit) return text;
    return text.length > limit ? text.slice(0, limit) : text;
  }

  function stripHtml(value) {
    return toString(value).replace(/<[^>]*>/g, "");
  }

  function getPath(obj, path) {
    if (!path || typeof path !== "string") return undefined;
    const segments = path.split(".").filter(Boolean);

    const resolve = (value, idx) => {
      if (idx >= segments.length) return value;
      if (value === undefined || value === null) return undefined;

      const segment = segments[idx];
      const isArray = segment.endsWith("[]");
      const key = isArray ? segment.slice(0, -2) : segment;
      const next = value && typeof value === "object" ? value[key] : undefined;

      if (isArray) {
        if (!Array.isArray(next)) return [];
        if (idx === segments.length - 1) return next;
        const mapped = next.map((item) => resolve(item, idx + 1));
        return Array.prototype.concat.apply([], mapped);
      }

      return resolve(next, idx + 1);
    };

    return resolve(obj, 0);
  }

  const helpers = Object.freeze({
    get: getPath,
    coalesce,
    join,
    truncate,
    stripHtml,
    toString,
  });

  function compileTransform(code) {
    if (!code || typeof code !== "string") {
      throw new Error("Transform code is missing.");
    }
    const wrapped = `"use strict";\n${code}\nreturn typeof transform === "function" ? transform : null;`;
    const factory = new Function(wrapped);
    const transform = factory();
    if (typeof transform !== "function") {
      throw new Error("Transform function not found.");
    }
    return transform;
  }

  self.onmessage = (event) => {
    const payload = event && event.data ? event.data : {};
    const requestId = payload.requestId;
    try {
      const transform = compileTransform(payload.code);
      const entries = Array.isArray(payload.entries) ? payload.entries : [];
      const rows = entries.map((entry) => {
        const output = transform(entry, helpers);
        if (output && typeof output === "object") return output;
        return {};
      });
      safePostMessage({ requestId, ok: true, rows });
    } catch (err) {
      safePostMessage({ requestId, ok: false, error: err && err.message ? err.message : String(err) });
    }
  };
})();
