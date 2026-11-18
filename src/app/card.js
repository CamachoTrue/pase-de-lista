import { STATUSES } from '../lib/excel-core.js';
import { confirmChange } from './dialogs.js';
import { refresh, renderSummary } from './main-view.js';
import { LABELS, dayRecords, group, save, tab, typeOf, ui } from './state.js';
import { $, escapeHtml, normalize, toast } from './util.js';

export function firstPending() {
  const g = group();
  const rec = dayRecords(false);
  return g.students.findIndex((s) => !rec[s.row]);
}
// Al abrir un grupo/fecha empieza en el primer alumno sin marcar.
export function startCard() {
  const i = firstPending();
  ui.cardIdx = i < 0 ? group().students.length : i;
}

export function renderCard(animate) {
  const g = group();
  const rec = dayRecords(false);
  const n = g.students.length;
  const marked = g.students.filter((s) => rec[s.row]).length;
  const finished = ui.cardIdx >= n;

  $('#progress-fill').style.width = `${(marked / n) * 100}%`;
  ['#card', '#big-choices', '.card-nav'].forEach((sel) => ($(sel).hidden = finished));
  $('#done').hidden = !finished;

  if (finished) {
    $('#card-pos').textContent = `${marked} de ${n} marcados`;
    const c = { A: 0, F: 0, R: 0, J: 0 };
    g.students.forEach((s) => rec[s.row] && c[rec[s.row]]++);
    $('#done-summary').innerHTML = STATUSES.map((k) => `<span class="${k}">${LABELS[k]}: ${c[k]}</span>`).join('');
    const pending = n - marked;
    $('#done-pending').textContent = pending ? `Quedaron ${pending} alumno(s) sin marcar.` : 'Todos los alumnos tienen marca.';
    $('#done-restart').textContent = pending ? 'Ir a los pendientes' : 'Volver al inicio';
    return;
  }

  const s = g.students[ui.cardIdx];
  const st = rec[s.row] || '';
  $('#card-pos').textContent = `Alumno ${ui.cardIdx + 1} de ${n}`;
  $('#card-no').textContent = s.no ? `No. ${s.no}` : '';
  $('#card-name').innerHTML = escapeHtml(s.name) + (s.added ? ' <span class="badge-new">NUEVO</span>' : '');
  $('#card-status').textContent = st ? `Marcado: ${LABELS[st]}` : '';
  const card = $('#card');
  card.className = `card ${st}`;
  if (animate) { void card.offsetWidth; card.classList.add('slide'); }
  document.querySelectorAll('.big-choice').forEach((b) => b.classList.toggle('on', b.dataset.s === st));
  $('#card-prev').disabled = ui.cardIdx === 0;
}


export async function markCard(status) {
  const g = group();
  if (ui.busy || ui.cardIdx >= g.students.length) return;
  const s = g.students[ui.cardIdx];
  const prev = dayRecords(false)[s.row];
  if (prev && prev !== status) {
    ui.busy = true;
    const ok = await confirmChange(s.name, prev, status);
    ui.busy = false;
    if (!ok) return; // se queda en el mismo alumno, sin cambios
    toast(`Cambiado a ${LABELS[status]}`);
  }
  const rec = dayRecords(true);
  rec[s.row] = status;
  save();
  document.querySelectorAll('.big-choice').forEach((b) => b.classList.toggle('on', b.dataset.s === status));
  if (navigator.vibrate) navigator.vibrate(15);
  ui.busy = true;
  setTimeout(() => {   // breve pausa para ver la selección y pasa al siguiente
    ui.busy = false;
    ui.cardIdx++;
    renderCard(true);
    renderSummary();
  }, 200);
}

$('#big-choices').addEventListener('click', (e) => {
  const b = e.target.closest('.big-choice');
  if (b) markCard(b.dataset.s);
});
$('#card-prev').addEventListener('click', () => { if (ui.cardIdx > 0) { ui.cardIdx--; renderCard(true); } });
$('#card-next').addEventListener('click', () => { ui.cardIdx++; renderCard(true); });
$('#done-review').addEventListener('click', () => { ui.view = 'list'; refresh(); });
$('#done-restart').addEventListener('click', () => {
  const i = firstPending();
  ui.cardIdx = i < 0 ? 0 : i;
  renderCard(true);
});
$('#to-list').addEventListener('click', () => { ui.view = 'list'; refresh(); });
$('#to-card').addEventListener('click', () => { ui.view = 'card'; startCard(); refresh(); });

// Teclado (en computadora): A/F/R/J para marcar, flechas para moverse.
document.addEventListener('keydown', (e) => {
  if (typeOf(tab()) !== 'asist' || ui.view !== 'card' || $('#screen-main').hidden || document.querySelector('dialog[open]')) return;
  if (/^(INPUT|SELECT|TEXTAREA)$/.test(e.target.tagName)) return;
  const k = e.key.toUpperCase();
  if (STATUSES.includes(k)) markCard(k);
  else if (e.key === 'ArrowLeft') $('#card-prev').click();
  else if (e.key === 'ArrowRight') $('#card-next').click();
});

export function renderStudents() {
  const g = group();
  const rec = dayRecords(false);
  const q = normalize($('#search').value);
  const list = $('#students');
  list.innerHTML = g.students
    .map((s, i) => [s, i])
    .filter(([s]) => !q || normalize(s.name).includes(q))
    .map(([s, i]) => {
      const st = rec[s.row] || '';
      const btns = STATUSES.map(
        (k) => `<button class="choice ${st === k ? 'on' : ''}" data-s="${k}" aria-pressed="${st === k}">${k}<small>${LABELS[k]}</small></button>`
      ).join('');
      return `<li class="student ${st}" data-row="${s.row}" data-idx="${i}">
        <div class="student-name"><span class="no">${escapeHtml(s.no)}</span><span>${escapeHtml(s.name)}${s.added ? '<span class="badge-new">NUEVO</span>' : ''}</span></div>
        <div class="choices">${btns}</div></li>`;
    })
    .join('');
  renderSummary();
}

$('#students').addEventListener('click', async (e) => {
  const nameEl = e.target.closest('.student-name');
  if (nameEl) {   // tocar el nombre: pasar lista uno por uno desde ese alumno
    ui.cardIdx = Number(nameEl.closest('.student').dataset.idx);
    ui.view = 'card';
    refresh();
    return;
  }
  const btn = e.target.closest('.choice');
  if (!btn) return;
  const li = btn.closest('.student');
  const row = li.dataset.row;
  const s = btn.dataset.s;
  const prev = dayRecords(false)[row];
  const next = prev === s ? undefined : s; // tocar la misma otra vez = quitar marca
  if (prev) {
    const name = group().students.find((x) => String(x.row) === row).name;
    if (!(await confirmChange(name, prev, next))) return;
  }
  const rec = dayRecords(true);
  if (next) rec[row] = next;
  else delete rec[row];
  save();
  li.className = `student ${rec[row] || ''}`;
  li.querySelectorAll('.choice').forEach((b) => {
    const on = b.dataset.s === rec[row];
    b.classList.toggle('on', on);
    b.setAttribute('aria-pressed', on);
  });
  renderSummary();
});
