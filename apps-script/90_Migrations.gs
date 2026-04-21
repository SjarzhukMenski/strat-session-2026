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
  // Заполните mapping ДО запуска — ФИО точно как в СправочникПроектов.Владелец проекта (C64)
  const FIO_TO_EMAIL = {
    // 'Иванов Иван Иванович': 'ivanov@example.com',
  };

  const sumSh = getMainSs().getSheetByName(SHEETS.PROJECTS_SUMMARY);
  const headers = sumSh.getRange(1, 1, 1, sumSh.getLastColumn()).getValues()[0];
  const colOwnerEmail = headers.indexOf('email владельца') + 1;
  if (colOwnerEmail === 0) throw new Error('Не найдена колонка "email владельца"');

  const dirSh = getMainSs().getSheetByName(SHEETS.DIR);
  const dirRows = readRows(dirSh);
  const sumRows = readRows(sumSh);

  sumRows.forEach(r => {
    const code = r['Лист'];
    const dir = dirRows.find(d => d['Код проекта (I2)'] === code);
    if (!dir) return;
    const fio = String(dir['Владелец проекта (C64)'] || '').trim();
    const email = FIO_TO_EMAIL[fio];
    if (email) {
      sumSh.getRange(r._row, colOwnerEmail).setValue(email);
      addAccess(email, 'owner', fio, 'migration');
    }
  });

  console.log('✓ migration_backfillOwnerEmails завершена');
}
