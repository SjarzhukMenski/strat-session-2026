function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    const session = validateSession(e.parameter.sessionToken);
    return jsonResponse(dispatch(action, 'GET', e.parameter, session));
  } catch (err) {
    return jsonResponse(fail(err.message || String(err), 500));
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = (e.parameter.action || body.action || '').trim();
    const session = validateSession(body.sessionToken);
    return jsonResponse(dispatch(action, 'POST', body, session));
  } catch (err) {
    return jsonResponse(fail(err.message || String(err), 500));
  }
}

function dispatch(action, method, params, session) {
  const lock = LockService.getDocumentLock();
  const isWrite = action.includes('create') || action.includes('update') ||
                  action.includes('setStatus') || action === 'auth.requestLink' ||
                  action === 'auth.activate';

  if (isWrite) lock.waitLock(10000);
  try {
    switch (action) {
      // Публичные
      case 'problems.list':  return problems_list(params);
      case 'problems.get':   return problems_get(params);
      case 'projects.list':  return projects_list(params);
      case 'projects.get':   return projects_get(params);

      // Auth
      case 'auth.requestLink': return auth_requestLink(params);
      case 'auth.activate':    return auth_activate(params);
      case 'auth.whoami':      return auth_whoami(session);

      // Запись
      case 'projects.create':    return projects_create(params, session);
      case 'projects.update':    return projects_update(params, session);
      case 'projects.setStatus': return projects_setStatus(params, session);
      case 'problems.setStatus': return problems_setStatus(params, session);

      default: return fail('unknown_action: ' + action, 404);
    }
  } finally {
    if (isWrite) lock.releaseLock();
  }
}

// --- Auth handlers ---

function auth_requestLink(params) {
  const email = String(params.email || '').trim().toLowerCase();
  if (!email) return fail('email required');
  try {
    if (lookupAccess(email)) {
      const token = createLinkToken(email);
      sendMagicLink(email, token, 'login');
    }
  } catch (e) {
    if (e.message !== 'rate_limited') throw e;
  }
  return ok({});
}

function auth_activate(params) {
  return activateToken(params.token, params.user_agent || '');
}

function auth_whoami(session) {
  if (!session) return fail('unauthorized', 401);
  const owned = getOwnedProjects(session.email);
  return ok({ email: session.email, role: session.role, name: session.name, ownedProjects: owned });
}

// --- Заглушки (будут реализованы в Phase 2+) ---

function problems_list()      { return ok([]); }
function problems_get()       { return ok({}); }
function projects_list()      { return ok([]); }
function projects_get()       { return ok({}); }
function projects_create()    { return fail('not_implemented'); }
function projects_update()    { return fail('not_implemented'); }
function projects_setStatus() { return fail('not_implemented'); }
function problems_setStatus() { return fail('not_implemented'); }
