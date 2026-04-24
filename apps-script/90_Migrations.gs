// Схема ручных ячеек карточки проекта (используется при создании, очистке шаблона, обновлении)
//
// I2  = формула =CONCATENATE("П";D2;"-";"К";F2) — НЕ трогаем
// D2  = номер проекта n (единственное что меняем для нового кода)
// F2  = "4" в шаблоне, остаётся константой
//
// Ручные ячейки (заполняются при создании / редактировании):
//   D2        — номер проекта
//   B7        — наименование проблемы
//   B17       — название проекта
//   B20       — описание решения
//   B25:H26   — метрики (диапазон)
//   E28       — срок достижения
//   E30       — тип проекта
//   H34       — (ставка, блок А)
//   H36       — средняя ставка
//   C40       — сумма блока Б
//   C44       — сумма блока В
//   C48       — сумма блока Г
//   C64       — владелец
//   B74:I78   — план действий (диапазон)
//
// Раздел 5 (заполняется только при статусе «Выполнен»):
//   B83:G85   — результаты (диапазон)
//   G88       — результат 1
//   G89       — результат 2
//   C95       — итоговый результат

function migration_prepareTemplate() {
  const ss = getMainSs();
  const tpl = ss.getSheetByName(SHEETS.TEMPLATE);
  if (!tpl) throw new Error('Лист _ШаблонКарточки не найден. Сначала создайте его как копию П7-К3.');

  const manualCells = [
    'D2',         // Номер проекта (I2 — формула, F2="4" — константа, не трогаем)
    'B7',         // Наименование проблемы
    'B17',        // Название проекта
    'B20',        // Описание решения
    'E28',        // Срок достижения
    'E30',        // Тип проекта
    'H34',        // Ставка блока А
    'H36',        // Средняя ставка
    'C40',        // Сумма блока Б
    'C44',        // Сумма блока В
    'C48',        // Сумма блока Г
    'C64',        // Владелец
  ];
  manualCells.forEach(a1 => tpl.getRange(a1).clearContent());

  // Диапазоны
  tpl.getRange('B25:H26').clearContent(); // Метрики
  tpl.getRange('B74:I78').clearContent(); // План действий

  // Раздел 5 (результаты)
  tpl.getRange('B83:G85').clearContent();
  tpl.getRange('G88').clearContent();
  tpl.getRange('G89').clearContent();
  tpl.getRange('C95').clearContent();

  console.log('✓ Шаблон подготовлен. Откройте лист _ШаблонКарточки и убедитесь визуально.');
}

function migration_backfillOwnerEmails() {
  // Читаем email из листа Доступ (Auth-таблица) по имени владельца.
  // Owners должны быть уже добавлены в Доступ вручную.
  const accessRows = readRows(getAuthSs().getSheetByName(AUTH_SHEETS.ACCESS));
  const nameToEmail = {};
  accessRows.forEach(r => {
    const name = String(r['имя'] || '').trim();
    if (name) nameToEmail[name] = String(r['email'] || '').trim().toLowerCase();
  });

  const sumSh = getMainSs().getSheetByName(SHEETS.PROJECTS_SUMMARY);
  const headers = sumSh.getRange(1, 1, 1, sumSh.getLastColumn()).getValues()[0];
  const colOwnerEmail = headers.indexOf('email владельца') + 1;
  if (colOwnerEmail === 0) throw new Error('Не найдена колонка "email владельца"');

  const dirSh = getMainSs().getSheetByName(SHEETS.DIR);
  const dirRows = readRows(dirSh);
  const sumRows = readRows(sumSh);

  let filled = 0;
  let missing = [];

  sumRows.forEach(r => {
    const code = r['Лист'];
    const dir = dirRows.find(d => d['Код проекта (I2)'] === code);
    if (!dir) return;
    const fio = String(dir['Владелец проекта (C64)'] || '').trim();
    const email = nameToEmail[fio];
    if (email) {
      sumSh.getRange(r._row, colOwnerEmail).setValue(email);
      filled++;
    } else if (fio) {
      missing.push(`${code}: "${fio}"`);
    }
  });

  console.log(`✓ migration_backfillOwnerEmails: заполнено ${filled}, не найдено в Доступ: ${missing.length}`);
  if (missing.length) console.log('Не найдены:', missing.join('; '));
}

// --- Система голосования ---

// Возвращает массив из count новых кодов P-XXX, начиная с max+1.
// Не пишет в sheet — вызывающий делает setValues атомарно.
function generateNextProblemCodes(sh, headers, count) {
  const codeIdx = headers.indexOf('Код');
  if (codeIdx < 0) return new Array(count).fill('');
  const lastRow = sh.getLastRow();
  let maxNum = 0;
  if (lastRow >= 2) {
    const codes = sh.getRange(2, codeIdx + 1, lastRow - 1, 1).getValues()
      .flat().map(c => String(c).trim()).filter(c => /^P-\d+$/.test(c));
    maxNum = codes.reduce((m, c) => {
      const n = parseInt(c.replace('P-', ''), 10);
      return n > m ? n : m;
    }, 0);
  }
  return Array.from({ length: count }, (_, i) =>
    'P-' + String(maxNum + i + 1).padStart(3, '0'));
}

// Вспомогательный хелпер: генерирует следующий Код вида P-001.
// Вызывается каждый раз при добавлении новой проблемы.
function generateProblemCode(sh, headers) {
  const codeIdx = headers.indexOf('Код');
  if (codeIdx < 0) return '';
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'P-001';
  const codes = sh.getRange(2, codeIdx + 1, lastRow - 1, 1).getValues()
    .flat().map(c => String(c).trim()).filter(c => /^P-\d+$/.test(c));
  const maxNum = codes.reduce((m, c) => {
    const n = parseInt(c.replace('P-', ''), 10);
    return n > m ? n : m;
  }, 0);
  return 'P-' + String(maxNum + 1).padStart(3, '0');
}

// Запустить один раз: заполняет Код для всех строк БазаПроблем где он пустой.
// ВАЖНО: перед запуском вручную добавьте колонку «Код» в БазаПроблем (можно в конец).
function migration_addProblemCodes() {
  const sh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const codeIdx = headers.indexOf('Код');
  if (codeIdx < 0) {
    console.log('STOP: колонка «Код» не найдена в БазаПроблем. Добавьте её вручную.');
    return;
  }
  const lastRow = sh.getLastRow();
  let filled = 0;
  for (let row = 2; row <= lastRow; row++) {
    const existing = String(sh.getRange(row, codeIdx + 1).getValue()).trim();
    if (existing) continue;
    const code = generateProblemCode(sh, headers);
    sh.getRange(row, codeIdx + 1).setValue(code);
    filled++;
  }
  console.log('migration_addProblemCodes: заполнено ' + filled + ' кодов.');
}

// Запустить один раз: создаёт лист «Голоса» в AUTH-таблице.
function migration_setupVotesSheet() {
  const ss = getAuthSs();
  if (ss.getSheetByName(AUTH_SHEETS.VOTES)) {
    console.log('Лист «Голоса» уже существует — пропускаем.');
    return;
  }
  const sh = ss.insertSheet(AUTH_SHEETS.VOTES);
  sh.getRange(1, 1, 1, 4).setValues([['email', 'problemCode', 'timestamp', 'учитывать_в_лимите']]);
  console.log('migration_setupVotesSheet: лист создан.');
}
