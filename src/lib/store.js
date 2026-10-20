// The API key lives in this browser profile's localStorage (per user, per machine), never in the
// document's settings, so it doesn't travel inside a .docx someone shares. Storage can be missing or
// throw (private mode, cleared site data, a locked-down WebView); the pane then works without a key.
const KEY = 'p3n.apiKey';
const FREE = 'p3n.freeRemaining';

function get(area, name) {
  try { return globalThis[area]?.getItem(name) ?? null; } catch { return null; }
}
function set(area, name, value) {
  try {
    if (value === null || value === '') globalThis[area]?.removeItem(name);
    else globalThis[area]?.setItem(name, value);
    return true;
  } catch { return false; }
}

export const apiKey = {
  get: () => get('localStorage', KEY) || '',
  set: (k) => set('localStorage', KEY, k),
  remove: () => set('localStorage', KEY, null),
};

// What X-Free-Remaining said last, for this session only.
export const freeLeft = {
  get: () => { const v = get('sessionStorage', FREE); return v === null ? null : Number(v); },
  set: (n) => set('sessionStorage', FREE, n === null ? null : String(n)),
};
