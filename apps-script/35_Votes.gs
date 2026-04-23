// votes_getState — публичный, возвращает счётчики + состояние текущего пользователя
function votes_getState(params, session) {
  const ss = getAuthSs();
  const sh = ss.getSheetByName(AUTH_SHEETS.VOTES);
  const total = CONFIG.VOTES_PER_USER;

  if (!sh || sh.getLastRow() < 2) {
    return ok({ counts: {}, myVotes: [], remaining: total, total });
  }

  const rows = readRows(sh);
  const active = rows.filter(r => r['учитывать_в_лимите'] === true);

  const counts = {};
  active.forEach(r => {
    const code = String(r['problemCode'] || '').trim();
    if (code) counts[code] = (counts[code] || 0) + 1;
  });

  let myVotes = [];
  let remaining = total;
  if (session) {
    const email = session.email.toLowerCase();
    const mine = active.filter(r => String(r['email']).toLowerCase() === email);
    myVotes = mine.map(r => String(r['problemCode']).trim());
    remaining = Math.max(0, total - mine.length);
  }

  return ok({ counts, myVotes, remaining, total });
}

// votes_toggle — требует сессии, добавляет или удаляет голос
function votes_toggle(params, session) {
  if (!session) return fail('unauthorized', 401);
  const problemCode = String(params.problemCode || '').trim();
  if (!problemCode) return fail('problemCode required');

  // Проверяем статус проблемы (только Парковка)
  const mainSh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const mainHeaders = mainSh.getRange(1, 1, 1, mainSh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const codeColIdx   = mainHeaders.indexOf('Код');
  const statusColIdx = mainHeaders.indexOf('Статус обработки');
  if (codeColIdx < 0) return fail('Код column not found in БазаПроблем');

  const lastMainRow = mainSh.getLastRow();
  if (lastMainRow < 2) return fail('problem_not_found', 404);
  const mainData = mainSh.getRange(2, 1, lastMainRow - 1, mainHeaders.length).getValues();
  const problemRow = mainData.find(r => String(r[codeColIdx]).trim() === problemCode);
  if (!problemRow) return fail('problem_not_found', 404);
  if (String(problemRow[statusColIdx]).trim() !== 'Парковка') {
    return fail('voting_not_allowed', 403);
  }

  const ss = getAuthSs();
  const sh = ss.getSheetByName(AUTH_SHEETS.VOTES);
  const rows = readRows(sh);
  const email = session.email.toLowerCase();

  // Ищем активный голос этого пользователя за эту проблему
  const existing = rows.find(r =>
    String(r['email']).toLowerCase() === email &&
    String(r['problemCode']).trim() === problemCode &&
    r['учитывать_в_лимите'] === true
  );

  if (existing) {
    // Отзываем голос — удаляем строку
    sh.deleteRow(existing._row);
  } else {
    // Проверяем лимит
    const myActive = rows.filter(r =>
      String(r['email']).toLowerCase() === email &&
      r['учитывать_в_лимите'] === true
    );
    if (myActive.length >= CONFIG.VOTES_PER_USER) return fail('vote_limit_reached', 403);
    // Добавляем голос
    sh.appendRow([session.email, problemCode, new Date(), true]);
  }

  return votes_getState(params, session);
}

// releaseVotes — вызывается из problems_setStatus при смене статуса с «Парковка»
function releaseVotes(problemCode) {
  const ss = getAuthSs();
  const sh = ss.getSheetByName(AUTH_SHEETS.VOTES);
  if (!sh || sh.getLastRow() < 2) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const codeColIdx  = headers.indexOf('problemCode');
  const limitColIdx = headers.indexOf('учитывать_в_лимите');
  if (codeColIdx < 0 || limitColIdx < 0) return;

  const lastRow = sh.getLastRow();
  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  data.forEach((row, i) => {
    if (String(row[codeColIdx]).trim() === problemCode && row[limitColIdx] === true) {
      sh.getRange(i + 2, limitColIdx + 1).setValue(false);
    }
  });
}

function test_votesGetState() {
  const r = votes_getState({}, null);
  if (!r.ok) throw new Error('votes_getState fail: ' + JSON.stringify(r));
  if (typeof r.data.counts !== 'object') throw new Error('counts missing');
  if (!Array.isArray(r.data.myVotes)) throw new Error('myVotes missing');
  console.log('✓ test_votesGetState', JSON.stringify(r.data));
}
