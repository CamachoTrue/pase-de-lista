// ---------- Pestaña Grupo ----------
// Todo lo que se configura del grupo en una pantalla: materia, criterios de evaluación (secciones y
// porcentajes), parciales, alumnos (altas y bajas), cuadernillo y ajustes que casi no se usan.
import { TYPES, sectionsOf, weightOf } from '../lib/excel-core.js';
import { confirmBulk } from './dialogs.js';
import { gname, group, save, state, todayISO, ui } from './state.js';
import { $, escapeHtml, normalize, toast } from './util.js';
import { paqueteGuardado, rolGuardado } from '../lib/nube.js';
import { abrir } from './menu.js';
import { refresh } from './main-view.js';
import { setSubject } from './subject.js';
import { openAddStudent, openRemoveStudent } from './students.js';
import { openSettings } from './settings.js';

const root = () => $('#vista-grupo');
const fechaC = (iso) => new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
let busca = '';

export function renderGrupo() {
  const g = group();
  const secs = sectionsOf(g);
  const total = secs.reduce((n, s) => n + weightOf(g.grading, s), 0);
  const per = state.periodos && state.periodos.parciales && state.periodos.parciales.length ? state.periodos : null;
  const escuela = paqueteGuardado() === 'empresarial';
  const director = rolGuardado() === 'director';
  const bajas = g.students.filter((s) => s.baja).length;
  root().innerHTML = `
    <div class="vista-head"><div><h2>Grupo</h2><p class="hint">Hoja del Excel: ${escapeHtml(g.sheet)} · ${g.students.length - bajas} alumnos${bajas ? ` (+${bajas} de baja)` : ''}</p></div></div>

    <section class="grupo-sec">
      <h3>Materia</h3>
      <form class="grupo-materia" data-form="materia">
        <input name="materia" maxlength="40" placeholder="Ej. Ética, Inglés II…" value="${escapeHtml(g.subject || '')}" aria-label="Materia">
        <button class="btn" type="submit">Guardar</button>
      </form>
      <p class="hint">Sirve si tienes el mismo grupo en dos materias, o para reconocerlo más fácil.</p>
    </section>

    <section class="grupo-sec">
      <h3>Criterios de evaluación</h3>
      <ul class="grupo-criterios">${secs.map((s) => `<li><span>${escapeHtml(s.name)}<small>${escapeHtml(TYPES[s.type].label)}</small></span><b>${weightOf(g.grading, s)} %</b></li>`).join('')}
        <li class="total ${total === 100 ? '' : 'mal'}"><span>Total</span><b>${total} %</b></li></ul>
      ${total === 100 ? '' : '<p class="hint mal-txt">Los porcentajes deben sumar 100 % para calcular calificaciones.</p>'}
      <button class="btn" data-g="criterios">Cambiar secciones y porcentajes</button>
      <p class="hint">Valen para el parcial en curso. Los parciales cerrados conservan las calificaciones con que se cerraron.</p>
    </section>

    <section class="grupo-sec">
      <h3>Parciales del semestre</h3>
      ${per ? `<ul class="grupo-criterios">${per.parciales.map((p) => `<li><span>Parcial ${p.numero}</span><b>${fechaC(p.inicio)} – ${fechaC(p.fin)}</b></li>`).join('')}</ul>` : '<p class="hint">Aún no hay fechas.</p>'}
      ${director ? '<p class="hint">Como director, las pones en Mi escuela → Parciales (valen para toda la escuela).</p>'
        : escuela ? `<p class="hint">Las fechas las pone tu director para toda la escuela${per ? '' : '; en cuanto las ponga aparecerán aquí'}.</p>`
        : `<button class="btn" data-g="fechas">${per ? 'Cambiar fechas' : 'Poner fechas'}</button>`}
    </section>

    <section class="grupo-sec">
      <h3>Alumnos</h3>
      <div class="row grupo-acc"><button class="btn" data-g="agregar">+ Agregar alumno</button>
        ${g.students.some((s) => s.added) ? '<button class="btn" data-g="quitar">Quitar uno agregado por error</button>' : ''}</div>
      <input type="search" class="grupo-busca" data-busca placeholder="Buscar alumno…" value="${escapeHtml(busca)}">
      <ul class="grupo-alumnos" data-alumnos></ul>
      <p class="hint">Dar de baja: desde hoy ya no aparece al pasar lista ni en los resúmenes. Lo capturado antes se conserva y lo puedes reactivar.</p>
    </section>

    <section class="grupo-sec">
      <h3>Cuadernillo</h3>
      <button class="btn" data-g="cuadernillo">Abrir el cuadernillo (PDF) del grupo</button>
    </section>

    <section class="grupo-sec">
      <h3>Más opciones</h3>
      <button class="btn" data-g="valores">Valores en el Excel (A, F, R, J, Sí, No)</button>
      <button class="btn danger-text" data-g="otro">Cargar otro archivo Excel</button>
    </section>`;
  pintarAlumnos();
}

function pintarAlumnos() {
  const g = group();
  const q = normalize(busca);
  const lista = g.students.filter((s) => !q || normalize(s.name).includes(q));
  root().querySelector('[data-alumnos]').innerHTML = lista.length ? lista.map((s) => `<li class="${s.baja ? 'baja' : ''}">
      <span class="nm"><span class="no">${escapeHtml(String(s.no ?? ''))}</span> ${escapeHtml(s.name)}
        ${s.baja ? `<small>De baja desde el ${fechaC(s.baja)}</small>` : s.added ? '<small>Agregado en la app</small>' : ''}</span>
      ${s.baja ? `<button class="link-btn" data-g="reactivar" data-row="${s.row}">Reactivar</button>`
        : `<button class="link-btn danger-text" data-g="baja" data-row="${s.row}">Dar de baja</button>`}</li>`).join('')
    : '<li class="empty">Ningún alumno coincide.</li>';
}

root().addEventListener('input', (e) => {
  if (!e.target.matches('[data-busca]')) return;
  busca = e.target.value;
  pintarAlumnos();
});
root().addEventListener('submit', (e) => {
  e.preventDefault();
  if (e.target.dataset.form === 'materia') { setSubject(e.target.elements.materia.value); renderGrupo(); }
});
root().addEventListener('click', async (e) => {
  const b = e.target.closest('[data-g]');
  if (!b) return;
  const g = group();
  const act = b.dataset.g;
  if (act === 'criterios' || act === 'fechas') abrir(() => import('./grades.js'), 'openGrades');
  if (act === 'agregar') openAddStudent();
  if (act === 'quitar') openRemoveStudent();
  if (act === 'cuadernillo') abrir(() => import('./book.js'), 'openBook');
  if (act === 'valores') openSettings();
  if (act === 'otro') {
    if (await confirmBulk('¿Cargar otro archivo Excel?', 'Se reemplazan las listas de este celular por las del archivo nuevo. Si solo cambia de parcial, no hace falta: usa «Cerrar parcial y pasar al siguiente» en Calificaciones. Exporta antes si necesitas un respaldo.', 'Sí, elegir archivo', true)) $('#file-input').click();
  }
  if (act === 'baja' || act === 'reactivar') {
    const s = g.students.find((x) => String(x.row) === b.dataset.row);
    if (!s) return;
    if (act === 'baja') {
      if (!(await confirmBulk(`¿Dar de baja a ${s.name}?`, `Desde hoy ya no aparece al pasar lista en ${gname(g)}. Lo que se capturó antes se conserva y lo puedes reactivar.`, 'Sí, dar de baja', true))) return;
      s.baja = todayISO();
    } else delete s.baja;
    if (save() === false) return; // sin acceso: no se guardó
    toast(act === 'baja' ? `${s.name} quedó de baja` : `${s.name} está activo otra vez`);
    renderGrupo();
  }
});
// Al cerrar las pantallas de criterios, alumnos o valores, se actualiza lo que se ve.
['grades', 'add-student', 'remove-student', 'settings', 'subject'].forEach((id) => {
  const d = $(`#${id}`);
  if (d) d.addEventListener('close', () => { if (ui.vista === 'grupo') { renderGrupo(); refresh(); } });
});
