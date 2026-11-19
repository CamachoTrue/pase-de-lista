// ---------- Dar de baja / reactivar alumnos (menú ⋮) ----------
// La baja vale desde hoy: el alumno ya no aparece al pasar lista ni en los resúmenes; lo capturado antes se conserva.
import { confirmBulk } from './dialogs.js';
import { gname, group, save, todayISO } from './state.js';
import { refresh } from './main-view.js';
import { $, escapeHtml, normalize, toast } from './util.js';

const fechaC = (iso) => new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });

export function openBajas() {
  $('#bajas-busca').value = '';
  pintar();
  $('#bajas').showModal();
}

function pintar() {
  const g = group();
  const q = normalize($('#bajas-busca').value);
  const lista = g.students.filter((s) => !q || normalize(s.name).includes(q));
  $('#bajas-lista').innerHTML = lista.length ? lista.map((s) => `<li class="${s.baja ? 'baja' : ''}">
      <span class="nm">${escapeHtml(s.name)}${s.baja ? `<small>De baja desde el ${fechaC(s.baja)}</small>` : ''}</span>
      ${s.baja ? `<button class="btn" data-b="reactivar" data-row="${s.row}">Reactivar</button>`
        : `<button class="btn danger-text" data-b="baja" data-row="${s.row}">Dar de baja</button>`}</li>`).join('')
    : '<li class="empty hint">Ningún alumno coincide.</li>';
}

$('#bajas-busca').addEventListener('input', pintar);
$('#bajas-lista').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-b]');
  if (!b) return;
  const g = group();
  const s = g.students.find((x) => String(x.row) === b.dataset.row);
  if (!s) return;
  if (b.dataset.b === 'baja') {
    if (!(await confirmBulk(`¿Dar de baja a ${s.name}?`, `Desde hoy ya no aparece al pasar lista en ${gname(g)}. Lo que se capturó antes se conserva y lo puedes reactivar.`, 'Sí, dar de baja', true))) return;
    s.baja = todayISO();
  } else delete s.baja;
  if (save() === false) { $('#bajas').close(); return; } // sin acceso: no se guardó
  toast(b.dataset.b === 'baja' ? `${s.name} quedó de baja` : `${s.name} está activo otra vez`);
  pintar();
  refresh();
});
