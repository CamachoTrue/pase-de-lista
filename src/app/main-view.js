import { STATUSES } from '../lib/excel-core.js';
import { renderCard, renderStudents, startCard } from './card.js';
import { renderCompact } from './compact.js';
import { OPTIONS, dayData, dayRecords, fmtDate, glabel, group, save, secs, state, tab, typeOf, ui, vistaApp } from './state.js';
import { rolGuardado } from '../lib/nube.js';
import { $, escapeHtml } from './util.js';

// ---------- Render ----------
// El director ve su escuela (no necesita grupos), salvo que elija "Mis grupos" porque también da clases.
export const vistaEscuela = () => ui.escuela === 'ejemplo' || (rolGuardado() === 'director' && vistaApp() !== 'grupos');

export function render() {
  const escuela = vistaEscuela();
  $('#screen-escuela').hidden = !escuela;
  $('#upload-escuela').hidden = rolGuardado() !== 'director';
  if (escuela) {
    $('#screen-upload').hidden = true;
    $('#screen-main').hidden = true;
    import('./escuela.js').then((m) => m.openEscuela()).catch((e) => console.error(e));
    return;
  }
  const has = !!(state && state.groups && state.groups.length);
  $('#screen-upload').hidden = has;
  $('#screen-main').hidden = !has;
  if (!has) return;

  const sel = $('#group-select');
  sel.innerHTML = state.groups
    .map((g, i) => `<option value="${i}">${escapeHtml(glabel(g))}</option>`)
    .join('');
  sel.value = state.ui.group;
  $('#date-input').value = state.ui.date;
  startCard();
  refresh();
}


export function refresh() {
  const opt = $('#group-select').options[state.ui.group];
  if (opt) opt.textContent = glabel(group());
  const t = tab();
  renderTabs();
  const att = typeOf(t) === 'asist';
  $('#card-view').hidden = !att || ui.view !== 'card';
  $('#list-view').hidden = !att || ui.view !== 'list';
  $('#compact-view').hidden = att;
  if (!att) renderCompact();
  else if (ui.view === 'card') renderCard(false);
  else renderStudents();
  renderSummary();
}

// ---------- Pestañas (una por sección del grupo) ----------
export function renderTabs() {
  const t = tab();
  const box = $('#tabs');
  const html = secs().map((x) => `<button class="tab ${x.id === t ? 'on' : ''}" role="tab" aria-selected="${x.id === t}" data-tab="${x.id}">${escapeHtml(x.name)}</button>`).join('');
  if (box.innerHTML !== html) {
    box.innerHTML = html;
    const on = box.querySelector('.tab.on');
    if (on) box.scrollLeft = Math.max(0, on.offsetLeft - box.clientWidth / 2 + on.offsetWidth / 2);
  }
}
$('#tabs').addEventListener('click', (e) => {
  const b = e.target.closest('.tab');
  if (!b || b.dataset.tab === tab()) return;
  state.ui.tab = b.dataset.tab;
  save();
  state.ui.work = 0;
  if (typeOf(state.ui.tab) === 'asist') startCard();
  refresh();
  window.scrollTo(0, 0);
});

export function renderSummary() {
  const g = group();
  const cat = tab();
  const t = typeOf(cat);
  if (t === 'num') {
    const day = dayData(cat, false);
    const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
    const avg = vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10 : null;
    $('#summary').innerHTML =
      `<span>${fmtDate(state.ui.date)}</span><span>${vals.length}/${g.students.length} capturados</span>` +
      (avg != null ? `<span>Promedio: ${avg}</span>` : '');
    return;
  }
  if (t === 'disc') {
    const day = dayData(cat, false);
    const zeros = g.students.filter((s) => day[s.row] === 0).length;
    $('#summary').innerHTML =
      `<span>${fmtDate(state.ui.date)}</span><span class="F">Con 0: ${zeros}</span><span>Los demás: 10</span>`;
    return;
  }
  if (t !== 'asist') {
    const day = dayData(cat, false);
    const vals = g.students.map((s) => day[s.row]).filter((v) => v != null);
    $('#summary').innerHTML =
      `<span>${vals.length}/${g.students.length} capturados</span>` +
      OPTIONS[t].map((o) => `<span class="${o.color}">${o.btn}: ${vals.filter((v) => v === o.v).length}</span>`).join('');
    return;
  }
  const rec = dayRecords(false);
  const counts = { A: 0, F: 0, R: 0, J: 0 };
  Object.values(rec).forEach((s) => counts[s]++);
  const marked = Object.keys(rec).length;
  $('#summary').innerHTML =
    `<span>${fmtDate(state.ui.date)}</span><span>${marked}/${g.students.length} marcados</span>` +
    STATUSES.map((k) => `<span class="${k}">${k}: ${counts[k]}</span>`).join('');
}
// Cambió el grupo o la fecha: vuelve a empezar desde el primer pendiente.
export function changed() {
  state.ui.work = 0; // al cambiar de día o grupo se empieza en el primer trabajo
  save();
  startCard();
  refresh();
}
$('#group-select').addEventListener('change', (e) => {
  state.ui.group = Number(e.target.value);
  $('#search').value = '';
  changed();
});
$('#date-input').addEventListener('change', (e) => {
  if (!e.target.value) return;
  state.ui.date = e.target.value;
  changed();
});
export function shiftDay(n) {
  const [y, m, d] = state.ui.date.split('-').map(Number);
  const dt = new Date(y, m - 1, d + n);
  state.ui.date = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
  $('#date-input').value = state.ui.date;
  changed();
}
$('#prev-day').addEventListener('click', () => shiftDay(-1));
$('#next-day').addEventListener('click', () => shiftDay(1));
$('#search').addEventListener('input', renderStudents);
