import { getSession, clearSession, apiGet } from './api.js';

export function renderNav() {
  const s = getSession();
  const nav = document.createElement('nav');
  nav.className = 'site-nav';
  nav.innerHTML = `
    <a href="board.html"><b style="white-space:nowrap;display:inline-flex;align-items:center;gap:4px">Мониторинг <img src="assets/idea.png" alt="" style="height:1.5em;filter:brightness(0)"></b></a>
    <a href="problems.html">Проблемы</a>
    <a href="map.html">Карта проектов</a>
    <a href="results.html">Результаты</a>
    <span class="nav-spacer"></span>
    ${s ? `<span id="nav-votes" class="nav-votes"></span><span>${s.email}</span> <button id="logout" class="button-secondary" style="padding:4px 10px;font-size:13px">Выйти</button>` : `<a href="login.html">Вход</a>`}
  `;
  document.body.prepend(nav);
  if (s) {
    document.getElementById('logout').onclick = () => { clearSession(); location.reload(); };
    loadNavVotes();
  }
}

export function renderNavVoteSlots(remaining, total) {
  const el = document.getElementById('nav-votes');
  if (!el) return;
  const used = total - remaining;
  const slots = Array.from({ length: total }, (_, i) =>
    `<span class="nav-vote-slot${i < used ? ' nav-vote-used' : ''}">🔥</span>`
  ).join('');
  el.innerHTML = slots;
}

async function loadNavVotes() {
  try {
    const res = await apiGet('votes.getState');
    if (res.ok) renderNavVoteSlots(res.data.remaining, res.data.total);
  } catch (_) {}
}
