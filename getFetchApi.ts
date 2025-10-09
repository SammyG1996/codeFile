// Utilis/getFetchApi.ts

export interface FetchApiInfo {
  spUrl: string;
  method?: 'GET' | 'POST' | 'DELETE' | 'PATCH' | 'MERGE';
  headers?: Record<string, string>;
  body?: any;
}

/* ───────────────────────────── CHANGED: begin (digest support) ───────────────────────────── */

// Minimal in-memory cache for request digests, keyed by site/web base URL.
type DigestEntry = { value: string; expiresOn: number };
const _digestCache = new Map<string, DigestEntry>();

// Find the base URL to call `/_api/contextinfo` on, from a full or relative SP REST URL.
function _getApiBase(spUrl: string): string {
  const i = spUrl.indexOf('/_api');
  if (i >= 0) return spUrl.slice(0, i);
  return typeof window !== 'undefined' ? window.location.origin : spUrl;
}

// Fetch a digest and cache it until the server’s timeout (with a tiny safety buffer).
async function _getDigest(baseUrl: string): Promise<string> {
  const now = Date.now();
  const cached = _digestCache.get(baseUrl);
  if (cached && cached.expiresOn > now + 5000) return cached.value;

  const r = await fetch(`${baseUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: { Accept: 'application/json;odata=nometadata' },
    credentials: 'same-origin'
  });
  if (!r.ok) throw new Error(`contextinfo failed: ${r.status}`);

  const j = await r.json();
  const value: string = j?.FormDigestValue;
  const ttlSec: number = Number(j?.FormDigestTimeoutSeconds ?? 0) || 900; // fallback 15m
  if (!value) throw new Error('contextinfo missing FormDigestValue');

  _digestCache.set(baseUrl, { value, expiresOn: now + (ttlSec - 15) * 1000 });
  return value;
}
/* ───────────────────────────── CHANGED: end (digest support) ───────────────────────────── */

export async function getFetchAPI(info: FetchApiInfo): Promise<any> {
  const method = info.method ?? 'GET';

  // Base headers (keep Accept as before)
  const headers: Record<string, string> = {
    Accept: 'application/json;odata=nometadata',
    ...(info.headers ?? {})
  };

  /* ───────────────────────────── CHANGED: add digest automatically ───────────────────────────── */
  if (method !== 'GET' && !headers['X-RequestDigest']) {
    const base = _getApiBase(info.spUrl);
    headers['X-RequestDigest'] = await _getDigest(base);
  }
  /* ───────────────────────────────────────────────────────────────────────────────────────────── */

  // Prepare body (unchanged behavior; only JSON.stringify plain objects if no Content-Type set)
  let body: BodyInit | undefined;
  const hasContentType = Object.keys(headers).some(h => h.toLowerCase() === 'content-type');
  if (info.body !== undefined && info.body !== null) {
    if (!hasContentType && typeof info.body === 'object' && !(info.body instanceof FormData)) {
      headers['Content-Type'] = 'application/json; charset=utf-8';
      body = JSON.stringify(info.body);
    } else {
      body = info.body as BodyInit;
    }
  }

  const resp = await fetch(info.spUrl, {
    method,
    headers,
    body,
    credentials: 'same-origin'
  });

  /* ───────────────────────────── CHANGED: safer response parsing ─────────────────────────────
     - Do NOT parse JSON for 204 (DELETE often returns No Content)
     - Avoid JSON parse when body is empty
  ------------------------------------------------------------------------------------------------ */
  if (resp.status === 204) return undefined;

  const contentType = (resp.headers.get('Content-Type') || '').toLowerCase();
  const isJson = contentType.includes('application/json');

  // Read once; some endpoints return 200 with empty body
  const text = await resp.text();
  if (!text) return undefined;

  if (isJson) {
    try { return JSON.parse(text); } catch { return text; } // fall back to raw text if malformed
  }
  /* ───────────────────────────────────────────────────────────────────────────────────────────── */

  // Non-JSON: return raw text (original behavior may have used json(); this is safer for binaries/empty)
  return text;
}