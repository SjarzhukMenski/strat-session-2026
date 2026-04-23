function problems_list(params) {
  const sh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const colMHeader = sh.getLastColumn() >= 13
    ? String(sh.getRange(1, 13, 1, 1).getValues()[0][0] || '').trim()
    : '';
  const rows = readRows(sh);
  let list = rows.map(r => {
    const raw = colMHeader ? r[colMHeader] : '';
    const hoursPerYear = (raw !== '' && raw != null) ? (parseInt(raw, 10) || null) : null;
    return {
      rowIndex: r._row,
      fio: r['1. ФИО'],
      department: r['2. Подразделение компании'],
      title: r['3. Название действия, операции или процесса (кратко, 5-7 слов)'],
      category: r['5. К какой категории относится эта операция?'],
      team: r['Команда'],
      hoursPerYear,
      status: r['Статус обработки'],
      projectCode: r['Проект'],
      responsible: r['ответственный'],
    };
  });

  const s = params.status;
  const dept = params.department;
  const team = params.team;
  const q = (params.q || '').toLowerCase();
  const uiStatus = params.uiStatus;

  if (s) list = list.filter(x => x.status === s);
  if (dept) list = list.filter(x => x.department === dept);
  if (team) list = list.filter(x => String(x.team) === String(team));
  if (q) list = list.filter(x =>
    (x.title || '').toLowerCase().includes(q) ||
    (x.fio || '').toLowerCase().includes(q)
  );

  if (uiStatus) {
    const projSh  = getMainSs().getSheetByName(SHEETS.MAP);
    const lastRow = projSh.getLastRow();
    const lastCol = projSh.getLastColumn();
    // Карта2026: строка 1 — агрегаты, строка 2 — заголовки, строки 3+ — данные
    const mapHeaders = lastRow >= 2
      ? projSh.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
      : [];
    const codeCol   = mapHeaders.indexOf('Лист');
    const statusCol = mapHeaders.indexOf('статус проекта');
    const projStatus = {};
    if (lastRow >= 3 && codeCol >= 0 && statusCol >= 0) {
      projSh.getRange(3, 1, lastRow - 2, lastCol).getValues().forEach(row => {
        if (row[codeCol]) projStatus[row[codeCol]] = row[statusCol] || '';
      });
    }

    list = list.filter(x => {
      if (x.status === 'Парковка') return uiStatus === 'park';
      if (x.status === 'Отменена') return uiStatus === 'cancelled';
      if (x.status === 'ОК' && x.projectCode) return uiStatus === 'inProject';
      return false;
    });

    // Прикрепляем статус проекта ко всем элементам вкладки inProject
    if (uiStatus === 'inProject') {
      list = list.map(x => ({
        ...x,
        projectStatus: projStatus[x.projectCode] || 'Не начат',
      }));
    }
  }

  // Фильтр по статусу проекта (клиентский или серверный)
  const ps = params.projectStatus;
  if (ps) list = list.filter(x => x.projectStatus === ps);

  return ok(list);
}

function problems_get(params) {
  const row = parseInt(params.row, 10);
  if (!row) return fail('row required');
  const sh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const values = sh.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = { _row: row };
  headers.forEach((h, i) => { obj[h] = values[i]; });
  return ok(obj);
}

function problems_setStatus(params, session) {
  if (!session || session.role !== 'coordinator') return fail('forbidden', 403);
  const row = parseInt(params.row, 10);
  const status = String(params.status || '').trim();
  if (!['Парковка', 'ОК', 'Отменена'].includes(status)) return fail('bad_status');

  const sh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colStatus    = headers.indexOf('Статус обработки') + 1;
  const colReason    = headers.indexOf('Причина отмены') + 1;
  const colChanged   = headers.indexOf('Изменён') + 1;
  const colChangedBy = headers.indexOf('Изменил') + 1;

  sh.getRange(row, colStatus).setValue(status);
  if (colReason > 0) sh.getRange(row, colReason).setValue(status === 'Отменена' ? (params.reason || '') : '');
  if (colChanged > 0) sh.getRange(row, colChanged).setValue(new Date());
  if (colChangedBy > 0) sh.getRange(row, colChangedBy).setValue(session.email);

  return ok({});
}

function test_problemsList() {
  const r = problems_list({});
  if (!r.ok || !Array.isArray(r.data)) throw new Error('fail');
  console.log('Всего:', r.data.length, 'первый:', JSON.stringify(r.data[0]));
  console.log('✓ test_problemsList');
}
