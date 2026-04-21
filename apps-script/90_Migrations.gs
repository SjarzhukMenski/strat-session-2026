function migration_prepareTemplate() {
  const ss = getMainSs();
  const tpl = ss.getSheetByName(SHEETS.TEMPLATE);
  if (!tpl) throw new Error('Лист _ШаблонКарточки не найден. Сначала создайте его как копию П7-К3.');

  const manualCells = [
    'I2',
    'B7',
    'B17',
    'B20',
    'B25', 'B26',
    'E28',
    'E30',
    'C64',
    'E38', 'E39',
    'E42', 'E43',
    'E46', 'E47',
    'C40', 'C44', 'C48',
    'H36',
  ];
  manualCells.forEach(a1 => tpl.getRange(a1).clearContent());

  tpl.getRange('B74:I78').clearContent();

  // TODO: уточнить у пользователя точные координаты раздела 5 (строки результатов)

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
