const API_BASE = 'https://script.google.com/macros/s/AKfycbxHZ-a4hPCEkDp1GLhqNnCCf8EI3CinwCOSqW5-beUclOBxsynKhUe1R-c_nItA7Cc/exec';

export async function apiGet(action, params = {}) {
  const url = new URL(API_BASE);
  url.searchParams.set('action', action);
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, v);
  const session = getSession();
  if (session) url.searchParams.set('sessionToken', session.sessionToken);
  const res = await fetch(url, { method: 'GET' });
  return res.json();
}

export async function apiPost(action, body = {}) {
  const session = getSession();
  if (session) body.sessionToken = session.sessionToken;
  const res = await fetch(`${API_BASE}?action=${action}`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
  });
  return res.json();
}

export function getSession() {
  try { return JSON.parse(localStorage.getItem('tm_session')); } catch { return null; }
}
export function setSession(s) { localStorage.setItem('tm_session', JSON.stringify(s)); }
export function clearSession() { localStorage.removeItem('tm_session'); }
