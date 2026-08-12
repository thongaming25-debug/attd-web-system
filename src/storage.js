// Polyfill for the window.storage API that Claude Artifacts provides.
// This lets the app run in a normal browser using localStorage instead.
// Swap this file out later if you plug in a real backend/database.

function fullKey(key, shared) {
  return `hrsuite:${shared ? "shared" : "local"}:${key}`;
}

window.storage = {
  async get(key, shared = false) {
    const raw = localStorage.getItem(fullKey(key, shared));
    if (raw === null) return null;
    return { key, value: raw, shared };
  },

  async set(key, value, shared = false) {
    localStorage.setItem(fullKey(key, shared), value);
    return { key, value, shared };
  },

  async delete(key, shared = false) {
    const existed = localStorage.getItem(fullKey(key, shared)) !== null;
    localStorage.removeItem(fullKey(key, shared));
    return { key, deleted: existed, shared };
  },

  async list(prefix = "", shared = false) {
    const marker = `hrsuite:${shared ? "shared" : "local"}:`;
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(marker + prefix)) {
        keys.push(k.slice(marker.length));
      }
    }
    return { keys, prefix, shared };
  },
};
