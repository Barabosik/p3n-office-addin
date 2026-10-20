// The p3n API (https://p3n.si/docs). Calls without a key are free, a few an hour per IP;
// a key from p3n.si/dashboard/keys goes in as a bearer token.
export const API = 'https://api.p3n.si';
export const SITE = 'https://p3n.si';
export const LINKS = {
  keys: `${SITE}/dashboard/keys`,
  billing: `${SITE}/dashboard/billing`,
  signup: `${SITE}/register`,
  docs: `${SITE}/docs`,
};
// The API's input limits: without a key, and with one (https://p3n.si/docs).
export const MAX_CHARS = { free: 5_000, key: 200_000 };

export class ApiError extends Error {
  constructor(status, type, message, extra = {}) {
    super(message);
    this.status = status;
    this.type = type;
    Object.assign(this, extra);
  }
}

const TYPE_BY_STATUS = { 401: 'unauthorized', 402: 'insufficient_credit', 403: 'forbidden', 404: 'not_found', 422: 'invalid_request', 429: 'rate_limited', 502: 'upstream_error' };

async function call(path, { method = 'GET', body, key, signal } = {}) {
  const headers = { Accept: 'application/json' };
  if (body) headers['Content-Type'] = 'application/json';
  if (key) headers.Authorization = `Bearer ${key}`;
  let res;
  try {
    res = await fetch(API + path, { method, headers, body: body ? JSON.stringify(body) : undefined, signal });
  } catch (e) {
    if (e.name === 'AbortError') throw e;
    throw new ApiError(0, 'network', "Can't reach p3n. Check your connection and try again.");
  }
  const free = res.headers.get('X-Free-Remaining');
  let data = null;
  try { data = await res.json(); } catch { /* empty or HTML error page */ }
  if (!res.ok) {
    const err = data?.error ?? {};
    throw new ApiError(res.status, err.type ?? TYPE_BY_STATUS[res.status] ?? 'error', err.message ?? data?.message ?? `p3n answered ${res.status}.`, {
      signupUrl: err.signup_url,
      retryAfter: Number(err.retry_after_seconds ?? res.headers.get('Retry-After')) || null,
      free: free === null ? null : Number(free),
    });
  }
  return { data, status: res.status, free: free === null ? null : Number(free) };
}

const sleep = (ms, signal) => new Promise((resolve, reject) => {
  const t = setTimeout(resolve, ms);
  signal?.addEventListener('abort', () => { clearTimeout(t); reject(new DOMException('Aborted', 'AbortError')); }, { once: true });
});

// Runs a skill on `text` and returns { run, free }. Long inputs come back as 202 and are polled.
export async function runSkill(skill, text, { key, options, signal } = {}) {
  const body = { text };
  if (options) body.options = options;
  let { data: run, status, free } = await call(`/v1/${skill}`, { method: 'POST', body, key, signal });
  for (let i = 0; (status === 202 || run?.status === 'queued' || run?.status === 'running') && i < 150; i++) {
    await sleep(2000, signal);
    ({ data: run, status } = await call(`/v1/runs/${run.id}`, { key, signal }));
  }
  if (run?.status === 'failed' || !run?.output) {
    throw new ApiError(502, 'run_failed', run?.error || 'The run failed. You were not charged. Try again.');
  }
  return { run, free };
}

// { email, balance_usd, key: { name, prefix } }
export async function me(key) {
  return (await call('/v1/me', { key })).data;
}
