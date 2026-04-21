import { getSession, clearSession } from './api.js';

export function renderNav() {
  const s = getSession();
  const nav = document.createElement('nav');
  nav.className = 'site-nav';
  nav.innerHTML = `
    <a href="board.html"><b>Мониторинг 2026</b></a>
    <a href="problems.html">Проблемы</a>
    <a href="map.html">Карта проектов</a>
    <span class="nav-spacer"></span>
    ${s ? `<span>${s.email}</span> <button id="logout" class="button-secondary" style="padding:4px 10px;font-size:13px">Выйти</button>` : `<a href="login.html">Вход</a>`}
  `;
  document.body.prepend(nav);
  if (s) document.getElementById('logout').onclick = () => { clearSession(); location.reload(); };
}
