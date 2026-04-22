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
    metric1:             g('B25'),
    metric2:             g('B26'),
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

function projects_create()    { return fail('not_implemented'); }
function projects_update()    { return fail('not_implemented'); }
function projects_setStatus() { return fail('not_implemented'); }
