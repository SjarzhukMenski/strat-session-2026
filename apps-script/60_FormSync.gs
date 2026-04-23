// Синхронизация ответов Google-формы → БазаПроблем
// Запускается ежедневно в 02:00 через триггер.
// Состояние хранится в Script Properties: FORM_LAST_ROW — номер последней
// обработанной строки в листе OperationsForm.

function syncFormResponses() {
  const RESPONSES_SS_ID = PropertiesService.getScriptProperties().getProperty('RESPONSES_SS_ID');
  if (!RESPONSES_SS_ID) throw new Error('RESPONSES_SS_ID не задан в Script Properties');

  const responsesSh = SpreadsheetApp.openById(RESPONSES_SS_ID).getSheetByName('OperationsForm');
  if (!responsesSh) throw new Error('Лист OperationsForm не найден в таблице ответов');

  const mainSh = getMainSs().getSheetByName(SHEETS.PROBLEMS);
  const props  = PropertiesService.getScriptProperties();

  const lastSyncedRow = parseInt(props.getProperty('FORM_LAST_ROW') || '1', 10);
  const totalRows     = responsesSh.getLastRow();

  if (totalRows <= lastSyncedRow) {
    console.log('syncFormResponses: нет новых записей (последняя строка: ' + totalRows + ')');
    return;
  }

  // Читаем заголовки обеих таблиц
  const respHeaders = responsesSh
    .getRange(1, 1, 1, responsesSh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());
  const mainHeaders = mainSh
    .getRange(1, 1, 1, mainSh.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const statusColIdx = mainHeaders.indexOf('Статус обработки');

  // Читаем только новые строки
  const startRow  = lastSyncedRow + 1;
  const numNew    = totalRows - lastSyncedRow;
  const newData   = responsesSh.getRange(startRow, 1, numNew, respHeaders.length).getValues();

  // Строим строки для БазаПроблем
  const rowsToAdd = newData.map(respRow => {
    const mainRow = new Array(mainHeaders.length).fill('');
    respHeaders.forEach((header, i) => {
      // "Отметка времени" — пропускаем (нет такого столбца в БазаПроблем)
      const mainIdx = mainHeaders.indexOf(header);
      if (mainIdx >= 0) mainRow[mainIdx] = respRow[i];
    });
    if (statusColIdx >= 0) mainRow[statusColIdx] = 'Парковка';
    return mainRow;
  });

  // Дописываем в конец БазаПроблем
  const appendAt = mainSh.getLastRow() + 1;
  mainSh.getRange(appendAt, 1, rowsToAdd.length, mainHeaders.length).setValues(rowsToAdd);

  // Сохраняем позицию
  props.setProperty('FORM_LAST_ROW', String(totalRows));

  console.log('syncFormResponses: добавлено ' + rowsToAdd.length +
    ' записей (строки ' + startRow + '–' + totalRows + ')');
}

// Вызвать один раз вручную из редактора Apps Script, чтобы установить триггер.
function setupFormSyncTrigger() {
  // Удаляем старые триггеры этой функции (защита от дублей)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncFormResponses')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncFormResponses')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  console.log('Триггер установлен: syncFormResponses ежедневно в 02:00');
}

// Вызвать один раз после первичного ручного заполнения БазаПроблем,
// чтобы сказать скрипту «всё что есть сейчас — уже добавлено, бери только новое».
function initFormSyncState() {
  const RESPONSES_SS_ID = PropertiesService.getScriptProperties().getProperty('RESPONSES_SS_ID');
  if (!RESPONSES_SS_ID) throw new Error('RESPONSES_SS_ID не задан в Script Properties');
  const sh = SpreadsheetApp.openById(RESPONSES_SS_ID).getSheetByName('OperationsForm');
  if (!sh) throw new Error('Лист OperationsForm не найден');
  const lastRow = sh.getLastRow();
  PropertiesService.getScriptProperties().setProperty('FORM_LAST_ROW', String(lastRow));
  console.log('initFormSyncState: FORM_LAST_ROW установлен в ' + lastRow + '. Новые записи будут добавляться начиная со строки ' + (lastRow + 1));
}

// Ручной запуск для разовой проверки / отладки.
function test_syncFormResponses() {
  syncFormResponses();
}
