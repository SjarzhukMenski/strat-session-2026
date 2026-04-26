function projects_list(params) {
  const mainSs = getMainSs();
  const mapSh  = mainSs.getSheetByName(SHEETS.MAP);
  const sumSh  = mainSs.getSheetByName(SHEETS.PROJECTS_SUMMARY);

  // Карта2026: строка 1 — агрегаты, строка 2 — заголовки, строки 3+ — данные
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 3) return ok([]);

  const mapHeaders = mapSh.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const mapData = mapSh.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const mapRows = mapData.map((r, i) => {
    const o = { _row: i + 3 };
    mapHeaders.forEach((h, j) => { o[h] = r[j]; });
    return o;
  });

  const sumRows = readRows(sumSh);
  const sumByCode = {};
  sumRows.forEach(r => { sumByCode[r['Лист']] = r; });

  const list = mapRows
    .filter(r => r['Лист'])
    .map(r => {
      const code = r['Лист'];
      const monthly = {};
      mapHeaders.forEach(h => {
        const m = String(h).match(/^(\d{4}-\d{2})$/);
        if (m) monthly[m[1]] = r[h];
      });
      const sum = sumByCode[code] || {};

      // Дополнительные поля из листа карточки проекта.
      // Читаем диапазон C30:H64 одним вызовом (35 строк × 6 столбцов).
      // Индексация: строка rowN → extra[rowN-30], столбец ColX → extra[...][colX-3]
      //   C=0, D=1, E=2, F=3, G=4, H=5
      //   H30 → [0][5]   budget
      //   H34 → [4][5]   hoursSaved
      //   E58 → [28][2]  paybackMonths
      //   C64 → [34][0]  owner
      let owner = '', budget = 0, hoursSaved = 0, paybackMonths = 0;
      const projSh = mainSs.getSheetByName(code);
      if (projSh) {
        const extra = projSh.getRange('C30:H64').getValues();
        budget        = Number(extra[0][5]  || 0);
        hoursSaved    = Number(extra[4][5]  || 0);
        paybackMonths = Number(extra[28][2] || 0);
        owner         = String(extra[34][0] || '');
      }

      return {
        code,
        name:         r['Название проекта'],
        type:         r['Сложность'],
        status:       r['статус проекта'] || 'Не начат',
        startMonth:   r['Месяц старта'],
        endMonth:     r['Ожидаемый месяц окончания'],
        team:         r['Команда (номер)'],
        ownerEmail:   sum['email владельца'] || '',
        owner,
        cancelReason: r['Причина отмены'] || '',
        monthlyBudget: monthly,
        totalBudget:  sum['Всего бюджет'] || 0,
        roi:          sum['Расчетный ROI'] || 0,
        effect:       sum['Нетто ээфект, руб.'] || 0,
        budget,
        hoursSaved,
        paybackMonths,
      };
    });

  let filtered = list;
  if (params.status) filtered = filtered.filter(x => x.status === params.status);
  if (params.team)   filtered = filtered.filter(x => String(x.team) === String(params.team));

  return ok(filtered);
}

function projects_get(params) {
  const code = params.code;
  if (!code) return fail('code required');
  const sh = getMainSs().getSheetByName(code);
  if (!sh) return fail('project_not_found', 404);

  const g = a1 => sh.getRange(a1).getValue();

  const data = {
    code,
    name:                g('B17'),
    team:                g('F2'),
    problemTitle:        g('B7'),
    problemDescription:  g('B10'),
    currentHours:        g('E12'),
    solutionDescription: g('B20'),
    metrics: (() => {
      // B(0)=name (ячейки B:E объединены), F(4)=ед.изм, G(5)=текущее, H(6)=целевое
      const raw = sh.getRange('B25:H26').getValues();
      return raw.filter(r => r[0]).map(r => ({
        name:    String(r[0] || ''),
        unit:    String(r[4] || ''),
        current: (r[5] !== '' && r[5] !== null && r[5] !== undefined) ? r[5] : null,
        target:  (r[6] !== '' && r[6] !== null && r[6] !== undefined) ? r[6] : null,
      }));
    })(),
    duration:            g('E28'),
    type:                g('E30'),
    budget:              g('H30'),
    months:              g('I30'),
    hoursSaved:          g('H34'),
    effectA:             g('C36'),
    averageRate:         g('H36'),
    effectB: { amount: g('C40'), desc: g('E38') },
    effectV: { amount: g('C44'), desc: g('E42') },
    effectG: { amount: g('C48'), desc: g('E46') },
    directEffect:   g('E50'),
    indirectEffect: g('E51'),
    totalEffect:    g('E55'),
    projectCost:    g('E56'),
    roi:            g('E57'),
    paybackMonths:  g('E58'),
    owner:          g('C64'),
    goal:           g('C66'),
    startMonth:     g('C70'),
    actions:        sh.getRange('B74:I78').getValues(),
  };

  // Статус из Карта2026
  const mapSh = getMainSs().getSheetByName(SHEETS.MAP);
  const lastCol = mapSh.getLastColumn();
  const headers = mapSh.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const codeCol   = headers.indexOf('Лист');
  const statusCol = headers.indexOf('статус проекта');
  const lastRow   = mapSh.getLastRow();
  if (lastRow >= 3 && codeCol >= 0 && statusCol >= 0) {
    const mapData = mapSh.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const mapRow  = mapData.find(r => r[codeCol] === code);
    data.status   = mapRow ? mapRow[statusCol] : 'Не начат';
  } else {
    data.status = 'Не начат';
  }

  return ok(data);
}

function projects_create(params, session) {
  if (!session) return fail('unauthorized', 401);
  if (session.role !== 'coordinator') return fail('forbidden', 403);

  const problemRow = parseInt(params.problemRow, 10);
  const name       = String(params.name       || '').trim();
  const ownerName  = String(params.ownerName  || '').trim();
  const ownerEmail = String(params.ownerEmail || '').trim().toLowerCase();
  const type       = String(params.type       || '').trim();
  const startMonth = String(params.startMonth || '').trim();

  if (!problemRow || !name || !ownerName || !ownerEmail ||
      !['S', 'M', 'L'].includes(type) || !/^\d{4}-\d{2}$/.test(startMonth)) {
    return fail('missing_or_invalid_params');
  }

  // Full labels matching data validation in the template sheet
  const TYPE_LABEL = {
    S: 'S (\u043c\u0430\u043b\u044b\u0439 \u043f\u0440\u043e\u0435\u043a\u0442)',
    M: 'M (\u0441\u0440\u0435\u0434\u043d\u0438\u0439 \u043f\u0440\u043e\u0435\u043a\u0442)',
    L: 'L (\u043a\u0440\u0443\u043f\u043d\u044b\u0439 \u043f\u0440\u043e\u0435\u043a\u0442)',
  };
  const typeLabel = TYPE_LABEL[type];

  const ss = getMainSs();

  // Find next project number n (max of existing Пn-К4 sheets + 1)
  const nums = ss.getSheets()
    .map(s => { const m = s.getName().match(/^П(\d+)-К4$/); return m ? parseInt(m[1], 10) : 0; })
    .filter(n => n > 0);
  const n    = nums.length > 0 ? Math.max(...nums) + 1 : 1;
  const code = '\u041f' + n + '-\u041a4';

  // Copy template sheet, rename, show
  const tpl = ss.getSheetByName(SHEETS.TEMPLATE);
  if (!tpl) return fail('template_not_found');
  const newSheet = tpl.copyTo(ss);
  newSheet.setName(code);
  newSheet.showSheet();

  // Fill manual cells
  newSheet.getRange('D2').setValue(n); // I2 formula auto-computes code П{n}-К4

  const probSh      = ss.getSheetByName(SHEETS.PROBLEMS);
  const probHeaders = probSh.getRange(1, 1, 1, probSh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const titleColIdx = probHeaders.findIndex(h => h.includes('\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u044f'));
  const problemTitle = titleColIdx >= 0
    ? String(probSh.getRange(problemRow, titleColIdx + 1).getValue() || '')
    : '';

  newSheet.getRange('B7').setValue(problemTitle);
  newSheet.getRange('B17').setValue(name);
  newSheet.getRange('E30').setValue(typeLabel);
  newSheet.getRange('C64').setValue(ownerName);

  // Compute endMonth from Служебный!J1:K3 (J=type S/M/L, K=duration months)
  const svcSh    = ss.getSheetByName(SHEETS.SERVICE);
  const typeData = svcSh.getRange('J1:K3').getValues();
  let durationMonths = 1;
  typeData.forEach(row => {
    if (String(row[0]).trim().toUpperCase().startsWith(type)) durationMonths = Number(row[1]) || durationMonths;
  });
  const endMonth = _addMonths(startMonth, durationMonths);

  // Read team from problem row
  const teamColIdx = probHeaders.indexOf('\u041a\u043e\u043c\u0430\u043d\u0434\u0430');
  const team = teamColIdx >= 0
    ? String(probSh.getRange(problemRow, teamColIdx + 1).getValue() || '')
    : '';

  // Append to СводПроекты (headers in row 1)
  const sumSh      = ss.getSheetByName(SHEETS.PROJECTS_SUMMARY);
  const sumHeaders = sumSh.getRange(1, 1, 1, sumSh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const sumRow     = sumHeaders.map(h => {
    switch (h) {
      case '\u041b\u0438\u0441\u0442':                           return code;
      case 'email \u0432\u043b\u0430\u0434\u0435\u043b\u044c\u0446\u0430':             return ownerEmail;
      case '\u041c\u0435\u0441\u044f\u0446 \u0441\u0442\u0430\u0440\u0442\u0430':                return _ymToDate(startMonth);
      case '\u041e\u0436\u0438\u0434\u0430\u0435\u043c\u044b\u0439 \u043c\u0435\u0441\u044f\u0446 \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f': return _ymToDate(endMonth);
      default: return '';
    }
  });
  sumSh.appendRow(sumRow);

  // Append to Карта2026 (headers in row 2, row 1 = aggregates)
  const mapSh      = ss.getSheetByName(SHEETS.MAP);
  const mapLastCol = mapSh.getLastColumn();
  const mapHeaders = mapSh.getRange(2, 1, 1, mapLastCol).getValues()[0].map(h => String(h).trim());
  const mapRow     = mapHeaders.map(h => {
    switch (h) {
      case '\u041b\u0438\u0441\u0442':                                  return code;
      case '\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u043f\u0440\u043e\u0435\u043a\u0442\u0430':                  return name;
      case '\u0421\u043b\u043e\u0436\u043d\u043e\u0441\u0442\u044c':                              return typeLabel;
      case '\u0441\u0442\u0430\u0442\u0443\u0441 \u043f\u0440\u043e\u0435\u043a\u0442\u0430':                  return '\u041d\u0435 \u043d\u0430\u0447\u0430\u0442';
      case '\u041c\u0435\u0441\u044f\u0446 \u0441\u0442\u0430\u0440\u0442\u0430':                       return _ymToDate(startMonth);
      case '\u041e\u0436\u0438\u0434\u0430\u0435\u043c\u044b\u0439 \u043c\u0435\u0441\u044f\u0446 \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f':      return _ymToDate(endMonth);
      case '\u041a\u043e\u043c\u0430\u043d\u0434\u0430 (\u043d\u043e\u043c\u0435\u0440)':                 return team;
      default: return '';
    }
  });
  mapSh.appendRow(mapRow);

  // Update БазаПроблем row
  const colStatus      = probHeaders.indexOf('\u0421\u0442\u0430\u0442\u0443\u0441 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438') + 1;
  const colProject     = probHeaders.indexOf('\u041f\u0440\u043e\u0435\u043a\u0442')           + 1;
  const colResponsible = probHeaders.indexOf('\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439')       + 1;
  if (colStatus > 0)      probSh.getRange(problemRow, colStatus).setValue('\u041e\u041a');
  if (colProject > 0)     probSh.getRange(problemRow, colProject).setValue(code);
  if (colResponsible > 0) probSh.getRange(problemRow, colResponsible).setValue(ownerName);

  // Grant whitelist access to owner (silently skips if already exists)
  addAccess(ownerEmail, 'owner', ownerName, session.email);

  // Protect new sheet like reference sheet П1-К1
  _protectSheetLikeReference(newSheet);

  // Rebuild project reference directory (function lives in the spreadsheet-bound script)
  try { createDetailedProjectsReference(); } catch (e) { console.warn('createDetailedProjectsReference:', e.message); }

  // Send welcome email with magic link
  try {
    const token       = createLinkToken(ownerEmail);
    const frontBase = CONFIG.WEBAPP_BASE_URL || '';
    const loginLink = frontBase
      ? frontBase + 'auth.html?token=' + token + '&next=' + encodeURIComponent('project.html?code=' + code)
      : '';
    const subject = (CONFIG.MAIL_SENDER_NAME || '\u041c\u043e\u043d\u0438\u0442\u043e\u0440\u0438\u043d\u0433') + ': \u0432\u044b \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d\u044b \u0432\u043b\u0430\u0434\u0435\u043b\u044c\u0446\u0435\u043c \u043f\u0440\u043e\u0435\u043a\u0442\u0430 ' + code;
    const body = '\u0417\u0434\u0440\u0430\u0432\u0441\u0442\u0432\u0443\u0439\u0442\u0435, ' + ownerName + '!\n\n' +
      '\u0412\u044b \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d\u044b \u0432\u043b\u0430\u0434\u0435\u043b\u044c\u0446\u0435\u043c \u043f\u0440\u043e\u0435\u043a\u0442\u0430 \u00ab' + name + '\u00bb (' + code + ').\n\n' +
      (loginLink
        ? '\u041f\u0435\u0440\u0435\u0439\u0434\u0438\u0442\u0435 \u043f\u043e \u0441\u0441\u044b\u043b\u043a\u0435 \u0434\u043b\u044f \u0432\u0445\u043e\u0434\u0430 \u0438 \u043f\u0440\u043e\u0441\u043c\u043e\u0442\u0440\u0430 \u043a\u0430\u0440\u0442\u043e\u0447\u043a\u0438:\n' + loginLink + '\n\n\u0421\u0441\u044b\u043b\u043a\u0430 \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u0442\u0435\u043b\u044c\u043d\u0430 30 \u043c\u0438\u043d\u0443\u0442.'
        : '\u0414\u043e\u0436\u0434\u0438\u0442\u0435\u0441\u044c \u0438\u043d\u0432\u0430\u0439\u0442\u0430 \u0434\u043b\u044f \u0432\u0445\u043e\u0434\u0430 \u0432 \u0441\u0438\u0441\u0442\u0435\u043c\u0443 \u043c\u043e\u043d\u0438\u0442\u043e\u0440\u0438\u043d\u0433\u0430.');
    MailApp.sendEmail({ to: ownerEmail, subject, body });
  } catch (e) {
    console.warn('welcome email failed:', e.message);
  }

  return ok({ code });
}

function _addMonths(ym, n) {
  const [y, mo] = ym.split('-').map(Number);
  const d = new Date(y, mo - 1 + n, 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function _ymToDate(ym) {
  const [y, mo] = ym.split('-').map(Number);
  return new Date(y, mo - 1, 1);
}

function _protectSheetLikeReference(sheet) {
  try {
    const refSh = getMainSs().getSheetByName(SHEETS.REFERENCE_P1K1);
    if (!refSh) return;
    const sourceProtections = refSh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (!sourceProtections.length) return;
    const sourceProt = sourceProtections[0];
    const unprotectedRanges = sourceProt.getUnprotectedRanges();
    const editors = sourceProt.getEditors().map(u => u.getEmail());

    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

    const newProt = sheet.protect();
    newProt.setDescription(sourceProt.getDescription() || 'Автозащита листа');
    newProt.removeEditors(newProt.getEditors());
    if (editors.length) newProt.addEditors(editors);
    if (sourceProt.isWarningOnly()) newProt.setWarningOnly(true);
    if (unprotectedRanges && unprotectedRanges.length) {
      newProt.setUnprotectedRanges(
        unprotectedRanges.map(r => sheet.getRange(r.getA1Notation()))
      );
    }
  } catch (e) {
    console.warn('_protectSheetLikeReference:', e.message);
  }
}

function meta_dropdowns() {
  const sh     = getMainSs().getSheetByName(SHEETS.SERVICE);
  const fio    = sh.getRange('R2:R18').getValues().flat().filter(Boolean).map(String);
  const months = sh.getRange('U2:U13').getValues().flat().filter(Boolean).map(d =>
    d instanceof Date
      ? d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')
      : String(d)
  );
  const typesRaw = sh.getRange('I1:K3').getValues();
  const types = typesRaw
    .filter(r => r[0])
    .map(r => ({ type: String(r[0]), budget: Number(r[1]) || 0, months: Number(r[2]) || 0 }));
  return ok({ fio, months, types });
}

function projects_update()    { return fail('not_implemented'); }
function projects_setStatus() { return fail('not_implemented'); }

function results_list(params, session) {
  if (!session) return fail('unauthorized', 401);

  const mainSs = getMainSs();
  const mapSh  = mainSs.getSheetByName(SHEETS.MAP);
  const probSh = mainSs.getSheetByName(SHEETS.PROBLEMS);

  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 3) return ok({ projects: [], summary: _resultsSummary([]) });

  const mapHeaders = mapSh.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const mapData    = mapSh.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const codeCol    = mapHeaders.indexOf('Лист');
  const statusCol  = mapHeaders.indexOf('статус проекта');

  const completedCodes = mapData
    .filter(r => String(r[statusCol]).trim() === 'Выполнен' && r[codeCol])
    .map(r => String(r[codeCol]));

  const probHeaders = probSh.getRange(1, 1, 1, probSh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const probData = probSh.getRange(2, 1, probSh.getLastRow() - 1, probSh.getLastColumn()).getValues();

  // Ищем по заголовку проблемы (B7 карточки) → колонка "3. Название действия..."
  // Подразделение — колонка "2. Подразделение компании" (регистронезависимо)
  const probTitleCol = probHeaders.findIndex(h => h.includes('Название действия'));
  const probDeptCol  = probHeaders.findIndex(h =>
    h.toLowerCase().includes('подразделение') && h.toLowerCase().includes('компании'));

  const projects = completedCodes.map(code => {
    const sh = mainSs.getSheetByName(code);
    if (!sh) return null;
    const g = a1 => sh.getRange(a1).getValue();

    // Метрики раздела 5: A=маркер(1=время), B=name, E=before, F=target, G=actual, H=deviation, I=effect
    // Диапазон A83:I85 — 9 столбцов (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8)
    const metricsRaw = sh.getRange('A83:I85').getValues();
    const metrics = metricsRaw
      .filter(r => r[1])
      .map(r => ({
        isTimeMetric: r[0] === 1 || r[0] === '1',
        name:         String(r[1] || ''),
        before:       r[4] !== '' ? Number(r[4]) : null,
        target:       r[5] !== '' ? Number(r[5]) : null,
        actual:       r[6] !== '' ? Number(r[6]) : null,
        deviation:    r[7] !== '' ? Number(r[7]) : null,
        effect:       r[8] !== '' ? Number(r[8]) : null,
      }));

    // Экономика E88:I90 — E=0, F=1(target), G=2(actual), H=3(deviation), I=4(pct)
    const econRaw = sh.getRange('E88:I90').getValues();
    const economics = {
      income: { target: econRaw[0][1], actual: econRaw[0][2], deviation: econRaw[0][3], pct: econRaw[0][4] },
      costs:  { target: econRaw[1][1], actual: econRaw[1][2], deviation: econRaw[1][3], pct: econRaw[1][4] },
      total:  { target: econRaw[2][1], actual: econRaw[2][2], deviation: econRaw[2][3], pct: econRaw[2][4] },
    };

    // Фактическая экономия времени: [до] − [фактически] (ч/год, всегда положительное)
    // Если маркер A=1 не проставлен — берём плановое значение из H34
    const timeMetric = metrics.find(m => m.isTimeMetric);
    const hoursSaved = (timeMetric && timeMetric.before != null && timeMetric.actual != null)
      ? Math.max(0, Number(timeMetric.before) - Number(timeMetric.actual))
      : Number(g('H34') || 0);
    const comment       = String(g('B93') || '');
    const completedRaw  = g('C95');
    const completedDate = completedRaw
      ? Utilities.formatDate(new Date(completedRaw), 'Europe/Moscow', 'yyyy-MM-dd')
      : '';

    // Подразделение: ищем строку в БазаПроблем по заголовку проблемы из B7
    const problemTitle = String(g('B7') || '').trim();
    let departments = [];
    if (probTitleCol >= 0 && probDeptCol >= 0 && problemTitle) {
      const probRow = probData.find(r => String(r[probTitleCol]).trim() === problemTitle);
      if (probRow) {
        departments = String(probRow[probDeptCol] || '')
          .split(/\s*\/\s*/)
          .map(s => s.trim())
          .filter(Boolean);
      }
    }

    return {
      code,
      name:          String(g('B17') || ''),
      owner:         String(g('C64') || ''),
      departments,
      completedDate,
      hoursSaved,
      metrics,
      economics,
      comment,
    };
  }).filter(Boolean);

  return ok({ projects, summary: _resultsSummary(projects) });
}

function _resultsSummary(projects) {
  const count             = projects.length;
  const totalHoursSaved   = projects.reduce((s, p) => s + (p.hoursSaved || 0), 0);
  const totalCostActual   = projects.reduce((s, p) => s + (Number(p.economics.costs.actual)  || 0), 0);
  const totalEffectTarget = projects.reduce((s, p) => s + (Number(p.economics.total.target)  || 0), 0);
  const totalEffectActual = projects.reduce((s, p) => s + (Number(p.economics.total.actual)  || 0), 0);
  const roiPct = totalCostActual > 0
    ? Math.round(totalEffectActual / totalCostActual * 1000) / 10
    : 0;
  return { count, totalHoursSaved, totalCostActual, totalEffectTarget, totalEffectActual, roiPct };
}
