// Логирование и дайджест неудачных попыток входа.
// Источник записей: auth_requestLink для unknown @santrade.by.
// Рассылка: hourly-триггер по будням (Пн-Пт, 09:00-17:59 Europe/Minsk),
//           тихо пропускает если новых строк нет.

function logFailedAttempt(email, userAgent) {
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.FAILED_LOGINS);
  if (!sh) return; // лист ещё не создан — silently skip
  const ua = String(userAgent || '').substring(0, 200); // обрезаем UA
  sh.appendRow([email, new Date(), ua]);
}

// Вызывается триггером. Читает с last_digest_row до getLastRow(),
// шлёт одно письмо координатору, обновляет Property.
function digestFailedLogins() {
  // Только будни
  const day = new Date().getDay(); // 0=Вс, 6=Сб
  if (day === 0 || day === 6) return;

  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.FAILED_LOGINS);
  if (!sh) return;

  const props = PropertiesService.getScriptProperties();
  const lastDigestRow = parseInt(props.getProperty('FAILED_LOGINS_LAST_ROW') || '1', 10);
  const currentLastRow = sh.getLastRow();

  if (currentLastRow <= lastDigestRow) return; // новых записей нет

  const numNew = currentLastRow - lastDigestRow;
  const startRow = lastDigestRow + 1;
  const newRows = sh.getRange(startRow, 1, numNew, 3).getValues();

  const accessRows = readRows(getAuthSs().getSheetByName(AUTH_SHEETS.ACCESS));
  const coord = accessRows.find(r => String(r['роль']).trim() === 'coordinator');
  if (!coord) return;

  const tz = 'Europe/Minsk';
  const fmt = d => Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd HH:mm');

  let body;
  if (numNew <= 10) {
    const lines = newRows.map(r =>
      `• ${fmt(r[1])} — ${r[0]}${r[2] ? ' (' + String(r[2]).substring(0, 80) + ')' : ''}`
    ).join('\n');
    body = `За последний час зафиксировано ${numNew} неудачн${numNew === 1 ? 'ая попытка' : 'ых попыток'} входа:\n\n${lines}\n\n— Мониторинг 2026`;
  } else {
    const freq = {};
    newRows.forEach(r => { freq[r[0]] = (freq[r[0]] || 0) + 1; });
    const top5 = Object.entries(freq).sort((a, b) => b[1] - a[1]).slice(0, 5)
      .map(([e, c]) => `  ${e} × ${c}`).join('\n');
    const first = fmt(newRows[0][1]);
    const last = fmt(newRows[numNew - 1][1]);
    body = `Подозрительная активность: ${numNew} попыток входа за последний час (${first} — ${last}).\n\nТоп-5 адресов:\n${top5}\n\nПолный список: лист «${AUTH_SHEETS.FAILED_LOGINS}» в TasksMonitoring_Auth.\n\n— Мониторинг 2026`;
  }

  MailApp.sendEmail({
    to: coord['email'],
    subject: `[Мониторинг] ${numNew} неудачн${numNew === 1 ? 'ая попытка' : 'ых попыток'} входа`,
    body,
    name: CONFIG.MAIL_SENDER_NAME,
  });

  props.setProperty('FAILED_LOGINS_LAST_ROW', String(currentLastRow));
  console.log(`digestFailedLogins: отправлено ${numNew} записей координатору ${coord['email']}`);
}

// Запустить один раз вручную из редактора.
function setupFailedLoginsTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'digestFailedLogins')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('digestFailedLogins')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('✓ Триггер установлен: digestFailedLogins раз в час');
}

// Запустить ОДИН раз после первого деплоя — зафиксировать baseline (все текущие строки считать обработанными).
function initFailedLoginsState() {
  const sh = getAuthSs().getSheetByName(AUTH_SHEETS.FAILED_LOGINS);
  if (!sh) throw new Error('Лист НеудачныеПопытки не найден');
  const lastRow = sh.getLastRow();
  PropertiesService.getScriptProperties().setProperty('FAILED_LOGINS_LAST_ROW', String(lastRow));
  console.log(`✓ FAILED_LOGINS_LAST_ROW = ${lastRow}. Уведомления пойдут о записях начиная со строки ${lastRow + 1}`);
}

function test_digestFailedLogins() {
  logFailedAttempt('test-fake@santrade.by', 'Mozilla/5.0 test');
  digestFailedLogins();
  console.log('✓ Проверьте инбокс координатора — должно прийти письмо с 1 записью');
}
