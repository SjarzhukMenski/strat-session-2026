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

  const access = lookupAccess(hit['email']);
  if (!access) return fail('not_in_whitelist', 403);

  const now = new Date();

  // Invalidate link token: пометить использованным И сдвинуть истекает в прошлое,
  // чтобы validateSession его больше не принял, даже если подставят его как sessionToken.
  // Колонки: 1=токен, 2=email, 3=создан, 4=истекает, 5=использован, 6=user_agent
  sh.getRange(hit._row, 4, 1, 3).setValues([[now, now, (userAgent || '') + ' [link]']]);

  // Create new session row с отдельным токеном
  const sessionToken = generateToken();
  const sessionExpires = new Date(now.getTime() + CONFIG.TOKEN_TTL_SESSION_HOURS * 3600000);
  sh.appendRow([sessionToken, hit['email'], now, sessionExpires, now, userAgent || '']);

  return ok({
    sessionToken,
    email: access.email,
    role: access.role,
    name: access.name,
    expiresAt: sessionExpires.toISOString()
  });
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
  const linkToken = createLinkToken('sergsantrade@gmail.com');
  console.log('Link token:', linkToken);

  const a = activateToken(linkToken, 'test-agent');
  if (!a.ok) throw new Error('activate failed: ' + JSON.stringify(a));

  // Ключевая проверка ротации: sessionToken ДОЛЖЕН отличаться от linkToken
  if (a.data.sessionToken === linkToken) throw new Error('token NOT rotated');

  // Валидация нового sessionToken — должна проходить
  const s = validateSession(a.data.sessionToken);
  if (!s || s.email !== 'sergsantrade@gmail.com') throw new Error('validate new token failed');

  // Валидация старого linkToken — должна НЕ проходить
  const stale = validateSession(linkToken);
  if (stale !== null) throw new Error('stale link token still valid!');

  console.log('✓ test_authFlow (rotation + invalidation)');
}
