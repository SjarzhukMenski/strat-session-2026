function lookupAccess(email) {
  const ss = getAuthSs();
  const sh = ss.getSheetByName(AUTH_SHEETS.ACCESS);
  const rows = readRows(sh);
  const e = String(email || '').trim().toLowerCase();
  const hit = rows.find(r => String(r['email']).trim().toLowerCase() === e);
  if (!hit) return null;
  return { email: hit['email'], role: hit['роль'], name: hit['имя'], _row: hit._row };
}

function addAccess(email, role, name, addedBy) {
  if (lookupAccess(email)) return;
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.ACCESS);
  sh.appendRow([email, role, name, new Date(), addedBy || 'system']);
}

function getOwnedProjects(email) {
  const sh = getMainSs().getSheetByName(SHEETS.PROJECTS_SUMMARY);
  const rows = readRows(sh);
  return rows
    .filter(r => String(r['email владельца'] || '').toLowerCase() === email.toLowerCase())
    .map(r => r['Лист']);
}

function test_lookupAccess() {
  const r = lookupAccess('sergsantrade@gmail.com');
  if (!r || r.role !== 'coordinator') throw new Error('lookup fail: ' + JSON.stringify(r));
  const miss = lookupAccess('nobody@example.com');
  if (miss !== null) throw new Error('miss fail');
  console.log('✓ test_lookupAccess');
}
