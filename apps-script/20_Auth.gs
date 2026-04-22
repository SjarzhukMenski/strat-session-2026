function createLinkToken(email) {
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.SESSIONS);
  const rows = readRows(sh);
  const windowMs = CONFIG.RATE_LIMIT_WINDOW_MIN * 60 * 1000;
  const recent = rows.filter(r =>
    String(r['email']).toLowerCase() === email.toLowerCase() &&
    (new Date() - new Date(r['создан'])) < windowMs
  );
  if (recent.length >= CONFIG.RATE_LIMIT_REQUESTS_PER_EMAIL)
    throw new Error('rate_limited');

  const token = generateToken();
  const created = new Date();
  const expires = new Date(created.getTime() + CONFIG.TOKEN_TTL_REQUEST_MIN * 60000);
  sh.appendRow([token, email, created, expires, '', '']);
  return token;
}

function activateToken(token, userAgent) {
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.SESSIONS);
  const rows = readRows(sh);
  const hit = rows.find(r => r['токен'] === token);
  if (!hit) return fail('invalid_token', 401);
  if (hit['использован']) return fail('already_used', 401);
  if (new Date() > new Date(hit['истекает'])) return fail('expired', 401);

  const now = new Date();
  const newExpires = new Date(now.getTime() + CONFIG.TOKEN_TTL_SESSION_HOURS * 3600000);
  sh.getRange(hit._row, 4).setValue(newExpires);
  sh.getRange(hit._row, 5).setValue(now);
  sh.getRange(hit._row, 6).setValue(userAgent || '');

  const access = lookupAccess(hit['email']);
  if (!access) return fail('not_in_whitelist', 403);
  return ok({ sessionToken: token, email: access.email, role: access.role, name: access.name, expiresAt: newExpires.toISOString() });
}

function validateSession(sessionToken) {
  if (!sessionToken) return null;
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.SESSIONS);
  const rows = readRows(sh);
  const hit = rows.find(r => r['токен'] === sessionToken);
  if (!hit || !hit['использован']) return null;
  if (new Date() > new Date(hit['истекает'])) return null;
  const access = lookupAccess(hit['email']);
  if (!access) return null;
  return { email: access.email, role: access.role, name: access.name };
}

function sendMagicLink(email, token, purpose) {
  const link = `${CONFIG.WEBAPP_BASE_URL}auth.html?token=${token}`;
  const subject = purpose === 'welcome'
    ? 'Вы назначены владельцем проекта — вход в систему'
    : 'Вход в мониторинг проектов';
  const body = `Ссылка активна 30 минут, не пересылайте её:\n\n${link}\n\n— Мониторинг 2026`;
  MailApp.sendEmail({ to: email, subject, body, name: CONFIG.MAIL_SENDER_NAME });
}

function cleanupOldSessions() {
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.SESSIONS);
  const rows = readRows(sh);
  const cutoff = new Date(Date.now() - 7 * 24 * 3600 * 1000);
  const toDelete = rows
    .filter(r => new Date(r['истекает']) < cutoff)
    .map(r => r._row)
    .sort((a, b) => b - a);
  toDelete.forEach(row => sh.deleteRow(row));
  console.log(`cleanupOldSessions: удалено ${toDelete.length}`);
}

function test_authFlow() {
  const t = createLinkToken('sergsantrade@gmail.com');
  console.log('Token:', t);
  const a = activateToken(t, 'test-agent');
  if (!a.ok) throw new Error('activate failed: ' + JSON.stringify(a));
  const s = validateSession(t);
  if (!s || s.email !== 'sergsantrade@gmail.com') throw new Error('validate failed');
  console.log('✓ test_authFlow');
}
