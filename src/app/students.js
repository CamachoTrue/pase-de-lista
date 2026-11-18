import { ASIST_ID, sectionsOf } from '../lib/excel-core.js';
import { refresh } from './main-view.js';
import { gname, group, save, state, ui } from './state.js';
import { $, escapeHtml, normalize, toast } from './util.js';

// ---------- Pasar lista ----------
// ---------- Quitar alumno (solo los agregados desde la app) ----------
// Está fuera del pase de lista a propósito: se busca al alumno, se elige y se confirma.
export let removeRow = null;

export function openRemoveStudent() {
  removeRow = null;
  $('#remove-search').value = '';
  renderRemoveList();
  $('#remove-student').showModal();
}

export function renderRemoveList() {
  const g = group();
  const q = normalize($('#remove-search').value);
  const added = g.students.filter((s) => s.added);
  const shown = added.filter((s) => !q || normalize(s.name).includes(q));
  $('#remove-list').innerHTML = !added.length
    ? '<li class="empty">No hay alumnos agregados en este grupo. Los alumnos que vienen del Excel original no se pueden quitar.</li>'
    : !shown.length
      ? '<li class="empty">Ningún alumno coincide con la búsqueda.</li>'
      : shown.map((s) => `<li data-row="${s.row}" class="${s.row === removeRow ? 'sel' : ''}"><span class="no">${escapeHtml(s.no)}</span>${escapeHtml(s.name)}</li>`).join('');
  const sel = g.students.find((s) => s.row === removeRow);
  const ok = $('#remove-ok');
  ok.disabled = !sel;
  ok.textContent = sel ? 'Quitar alumno elegido' : 'Elige un alumno';
}

$('#remove-search').addEventListener('input', renderRemoveList);
$('#remove-list').addEventListener('click', (e) => {
  const li = e.target.closest('li[data-row]');
  if (!li) return;
  removeRow = Number(li.dataset.row);
  renderRemoveList();
});
$('#remove-ok').addEventListener('click', () => {
  const g = group();
  const idx = g.students.findIndex((s) => s.row === removeRow && s.added);
  if (idx < 0) return;
  const s = g.students[idx];
  g.students.splice(idx, 1);
  sectionsOf(g).forEach((x) => Object.values(g[x.id] || {}).forEach((day) => {
    if (Array.isArray(day)) day.forEach((w) => delete w.vals[s.row]); // trabajos
    else delete day[s.row];
  }));
  save();
  if (ui.cardIdx > g.students.length) ui.cardIdx = g.students.length;
  $('#remove-student').close();
  refresh();
  toast(`${s.name} quitado de ${gname(g)}`);
});

// ---------- Agregar alumno ----------

export function openAddStudent() {
  const g = group();
  const nums = g.students.map((s) => Number(s.no)).filter(Number.isFinite);
  $('#add-group').textContent = `Grupo ${gname(g)} · ${g.students.length} alumnos`;
  $('#add-name').value = '';
  $('#add-no').value = nums.length ? Math.max(...nums) + 1 : g.students.length + 1;
  $('#add-error').textContent = '';
  const d = $('#add-student');
  d.returnValue = '';
  d.showModal();
  $('#add-name').focus();
}

$('#add-form').addEventListener('submit', (e) => {
  if (e.submitter && e.submitter.value === 'cancel') return; // cierra sin agregar
  e.preventDefault();
  const g = group();
  const name = $('#add-name').value.replace(/\s+/g, ' ').trim().toUpperCase();
  const no = $('#add-no').value.trim();
  if (!name) { $('#add-error').textContent = 'Escribe el nombre del alumno.'; return; }
  if (g.students.some((s) => normalize(s.name) === normalize(name))) {
    $('#add-error').textContent = 'Ya hay un alumno con ese nombre en este grupo.';
    return;
  }
  // Va en la siguiente fila libre debajo del último alumno.
  const row = Math.max(g.headerRow, ...g.students.map((s) => s.row)) + 1;
  g.students.push({ row, no, name, added: true });
  save();
  $('#add-student').close();
  // Lo deja en pantalla para pasarle lista de inmediato.
  state.ui.tab = ASIST_ID;
  save();
  ui.view = 'card';
  ui.cardIdx = g.students.length - 1;
  refresh();
  toast(`${name} agregado a ${gname(g)}`);
});
