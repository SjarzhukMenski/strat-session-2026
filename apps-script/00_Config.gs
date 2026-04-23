const CONFIG = {
  MAIN_SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  AUTH_SHEET_ID: PropertiesService.getScriptProperties().getProperty('AUTH_SHEET_ID'),
  WEBAPP_BASE_URL: PropertiesService.getScriptProperties().getProperty('WEBAPP_BASE_URL'),
  MAIL_SENDER_NAME: PropertiesService.getScriptProperties().getProperty('MAIL_SENDER_NAME') || 'Мониторинг',
  TOKEN_TTL_REQUEST_MIN: 30,
  TOKEN_TTL_SESSION_HOURS: 24,
  RATE_LIMIT_REQUESTS_PER_EMAIL: 3,
  RATE_LIMIT_WINDOW_MIN: 10,
  VOTES_PER_USER: parseInt(PropertiesService.getScriptProperties().getProperty('VOTES_PER_USER') || '3', 10),
};

const SHEETS = {
  PROBLEMS: 'БазаПроблем',
  PROJECTS_SUMMARY: 'СводПроекты',
  MAP: 'Карта2026',
  DIR: 'СправочникПроектов',
  SERVICE: 'Служебный',
  TEMPLATE: '_ШаблонКарточки',
  REFERENCE_P1K1: 'П1-К1',
};

const AUTH_SHEETS = { ACCESS: 'Доступ', SESSIONS: 'Сессии', VOTES: 'Голоса' };
