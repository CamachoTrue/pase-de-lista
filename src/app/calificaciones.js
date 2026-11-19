// ---------- Pestaña Calificaciones ----------
// Todo lo de calificar en una pantalla: el resumen del grupo (semáforo) por periodo — todo, cada
// parcial o el semestre —, el parcial en curso con «Cerrar y pasar al siguiente», y exportar.
// Tocar un alumno abre su resumen individual.
import { SEMESTRAL, activos, enPeriodo, finalGrade, groupReport, parcialDe, periodosListos, round1, semestral } from '../lib/excel-core.js';
import { cerrar, cierreActivo, fechaCorta, reabrir } from './cierres.js';
import { confirmBulk } from './dialogs.js';
import { gname, group, state, todayISO, ui } from './state.js';
import { $, escapeHtml, toast } from './util.js';
import { paqueteGuardado } from '../lib/nube.js';
import { abrir, irA } from './menu.js';
import { openExportar, openPreview } from './export.js';

const ST = { rojo: 'En riesgo', amarillo: 'Observación', verde: 'Bien', none: 'Sin datos' };
let periodo = null; // 'todo' | número de parcial | 'sem'
let filtro = 'all';
const root = () => $('#vista-calif');
const pars = () => (periodosListos(state.periodos) ? state.periodos.parciales : []);
const fechaC = (iso) => new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
// Parcial en curso: el de hoy; si hoy no cae en ninguno, el primero sin cerrar.
function enCurso(g) {
  const hoy = parcialDe(state.periodos, todayISO());
  if (hoy) return hoy;
  const abierto = pars().find((p) => !cierreActivo(g, 'parcial', p.numero));
  return abierto ? abierto.numero : null;
}

export function renderCalif() {
  const g = group();
  if (periodo == null) periodo = enCurso(g) || 'todo';
  if (typeof periodo === 'number' && !pars().some((p) => p.numero === periodo)) periodo = 'todo';
  const chips = pars().length
    ? `<nav class="calif-periodos">${[['todo', 'Todo'], ...pars().map((p) => [p.numero, `Parcial ${p.numero}`]), ['sem', 'Semestral']]
      .map(([k, t]) => `<button class="per-tab ${periodo === k ? 'on' : ''}" data-periodo="${k}">${t}</button>`).join('')}</nav>` : '';
  root().innerHTML = `
    <div class="vista-head"><div><h2>Calificaciones</h2><p class="hint">${escapeHtml(gname(g))}</p></div>
      <div class="vista-acc"><button class="btn" data-c="previa">Vista previa</button><button class="btn primary" data-c="exportar">Exportar</button></div></div>
    ${chips}
    ${tarjetaParcial(g)}
    <div data-cuerpo></div>`;
  if (periodo === 'sem') pintarSemestral(g); else pintarResumen(g);
}

// Estado del parcial y la acción principal: cerrar y pasar al siguiente.
function tarjetaParcial(g) {
  if (!pars().length) {
    const escuela = paqueteGuardado() === 'empresarial';
    return `<section class="calif-card"><b>Sin parciales</b>
      <p class="hint">${escuela ? 'Tu director aún no pone las fechas de los parciales. Mientras, aquí ves la calificación de todo lo capturado.'
        : 'Pon las fechas de tus parciales para ver la calificación de cada uno y cerrarlos.'}</p>
      ${escuela ? '' : '<button class="btn" data-c="fechas">Poner fechas de parciales</button>'}</section>`;
  }
  const n = typeof periodo === 'number' ? periodo : periodo === 'todo' ? enCurso(g) : null;
  if (n == null) return '';
  const p = pars().find((x) => x.numero === n);
  const c = cierreActivo(g, 'parcial', n);
  const ultimo = n === pars().at(-1).numero;
  if (c) {
    return `<section class="calif-card cerrado"><b>Parcial ${n} cerrado</b>
      <p class="hint">Se cerró el ${fechaCorta(c.cerrado_en)} · sus calificaciones ya son definitivas.</p>
      <button class="link-btn" data-c="reabrir" data-n="${n}">Reabrir el parcial ${n}</button></section>`;
  }
  const hoy = todayISO();
  if (p.inicio > hoy) {
    return `<section class="calif-card cerrado"><b>Parcial ${n} aún no empieza</b>
      <p class="hint">Empieza el ${fechaC(p.inicio)}. Lo que captures desde ese día cuenta para este parcial.</p></section>`;
  }
  const estado = p.fin < hoy ? `terminó el ${fechaC(p.fin)} · falta cerrarlo` : `en curso · del ${fechaC(p.inicio)} al ${fechaC(p.fin)}`;
  return `<section class="calif-card"><b>Parcial ${n}</b> <span class="hint">· ${estado}</span>
    <p class="hint">${ultimo ? 'Al terminar, ciérralo para dejar sus calificaciones como definitivas.'
      : `Al terminar, ciérralo: sus calificaciones quedan guardadas y pasas al parcial ${n + 1}, donde puedes cambiar porcentajes, secciones y alumnos.`}</p>
    <button class="btn primary" data-c="cerrar" data-n="${n}">${ultimo ? `Cerrar el parcial ${n}` : `Cerrar parcial ${n} y pasar al ${n + 1}`}</button></section>`;
}

function pintarResumen(g) {
  const p = pars().find((x) => x.numero === periodo);
  const vista = p ? enPeriodo(g, p.inicio, p.fin) : g;
  const c = p ? cierreActivo(g, 'parcial', p.numero) : null;
  const r = groupReport(vista, (c && c.criterios) || g.grading);
  // Un parcial cerrado muestra la calificación con que se cerró.
  if (c) r.rows.forEach((x) => { const v = c.calificaciones[x.student.row]; if (v) x.final = v.final; });
  const st = r.stats;
  const tile = (v, label, cls = '') => `<div class="gr-tile ${cls}"><b>${v}</b><span>${label}</span></div>`;
  const chip = (k, n, label) => `<button class="gr-chip ${k} ${filtro === k ? 'on' : ''}" data-f="${k}">${label} <b>${n}</b></button>`;
  const order = { rojo: 0, amarillo: 1, verde: 2, none: 3 };
  const rows = r.rows.filter((x) => filtro === 'all' || x.status === filtro)
    .sort((a, b) => order[a.status] - order[b.status] || (a.final ?? 99) - (b.final ?? 99) || (a.att ?? 101) - (b.att ?? 101));
  root().querySelector('[data-cuerpo]').innerHTML = `
    <div class="gr-tiles">${tile(st.alumnos, 'Alumnos')}${tile(st.asistencia == null ? '—' : `${round1(st.asistencia)} %`, 'Asistencia')}
      ${tile(st.promedio == null ? '—' : st.promedio, r.graded ? `Promedio (sobre ${r.escala})` : 'Promedio (faltan %)')}${tile(st.rojo, 'En riesgo', st.rojo ? 'rojo' : '')}</div>
    ${r.graded ? '' : '<p class="hint">Para calcular calificaciones, pon cuánto vale cada sección en la pestaña <b>Grupo</b>. <button class="link-btn" data-c="criterios">Ponerlos ahora ›</button></p>'}
    <div class="gr-chips">${chip('all', st.alumnos, 'Todos')}${chip('rojo', st.rojo, 'Rojo')}${chip('amarillo', st.amarillo, 'Amarillo')}${chip('verde', st.verde, 'Verde')}${st.none ? chip('none', st.none, 'Sin datos') : ''}</div>
    <details class="gr-criteria"><summary>¿Cómo se decide el color?</summary><ul>${r.criteria.map((t) => `<li>${escapeHtml(t)}</li>`).join('')}</ul></details>
    <p class="hint calif-tip">Toca un alumno para ver todo su detalle. <button class="link-btn" data-c="buscar">Buscar alumno ›</button></p>
    <ul class="gr-list calif-lista">${rows.length ? rows.map((x) => `<li data-row="${x.student.row}" class="${x.status}">
        <span class="gr-dot" title="${ST[x.status]}"></span>
        <span class="gr-name"><b>${escapeHtml(x.student.name)}</b><small>${x.reasons.length ? escapeHtml(x.reasons.join(' · ')) : ST[x.status]}</small></span>
        <span class="gr-num"><b>${x.final ?? '—'}</b><small>calif.</small></span>
        <span class="gr-num"><b>${x.att == null ? '—' : `${Math.round(x.att)}%`}</b><small>asist.</small></span>
        <span class="calif-ir" aria-hidden="true">›</span></li>`).join('') : '<li class="empty">No hay alumnos en este color.</li>'}</ul>`;
}

function pintarSemestral(g) {
  const gr = g.grading;
  const cfg = Object.assign({}, SEMESTRAL, g.semestral);
  const filas = activos(g).map((s) => [s, semestral(g, s.row, gr, state.periodos, cfg)]);
  const np = pars().length;
  const aprobados = filas.filter(([, r]) => r.aprobado === true).length;
  const reprobados = filas.filter(([, r]) => r.aprobado === false).length;
  const cs = cierreActivo(g, 'semestre');
  root().querySelector('[data-cuerpo]').innerHTML = `
    <div class="gr-tiles"><div class="gr-tile"><b>${aprobados}</b><span>Aprueban</span></div><div class="gr-tile ${reprobados ? 'rojo' : ''}"><b>${reprobados}</b><span>Reprueban</span></div></div>
    <p class="hint">${cfg.modo === 'final' ? `Parciales ${cfg.pesoParciales} % + examen final ${cfg.pesoFinal} %` : 'Promedio de los parciales'}${cs ? ` · semestre cerrado el ${fechaCorta(cs.cerrado_en)}` : ''}.
      <button class="link-btn" data-c="semestral">Cambiar cómo se calcula, capturar el examen final o cerrar el semestre ›</button></p>
    <div class="preview-scroll"><table class="sheet calif-sem"><thead><tr><th>Alumno</th>${Array.from({ length: np }, (_, i) => `<th>P${i + 1}</th>`).join('')}<th>Semestral</th></tr></thead><tbody>
    ${filas.map(([s, r]) => {
      const v = cs && cs.calificaciones[s.row] ? cs.calificaciones[s.row].reportada : r.reportada;
      return `<tr data-row="${s.row}" class="dir-click"><td class="name">${escapeHtml(s.name)}</td>${r.parciales.map((x) => `<td class="v">${x ?? '—'}</td>`).join('')}
        <td class="v ${r.aprobado === false ? 'dir-rojo' : ''}"><b>${v ?? '—'}</b></td></tr>`;
    }).join('')}</tbody></table></div>`;
}

// ---------- Cerrar el parcial y pasar al siguiente ----------
async function cerrarParcial(n) {
  const g = group();
  const p = pars().find((x) => x.numero === n);
  const sig = pars().find((x) => x.numero === n + 1);
  if (!(await confirmBulk(`¿Cerrar el parcial ${n} de ${gname(g)}?`,
    `Sus calificaciones quedan guardadas como definitivas: las puedes ver y exportar cuando quieras, y reabrirlo si hay un error.${sig ? ` Después pasas al parcial ${n + 1}.` : ''}`,
    'Sí, cerrar'))) return;
  const vista = enPeriodo(g, p.inicio, p.fin);
  const cal = {};
  g.students.forEach((s) => { const r = finalGrade(vista, s.row, g.grading); cal[s.row] = { final: r.final, partes: r.parts }; });
  cerrar(g, 'parcial', n, cal, g.grading);
  if (!cierreActivo(g, 'parcial', n)) return; // sin acceso: no se guardó
  toast(`Parcial ${n} cerrado`);
  if (!sig) { renderCalif(); return; }
  periodo = n + 1;
  renderCalif();
  // Lo que suele cambiar al empezar otro parcial.
  root().querySelector('.calif-card').insertAdjacentHTML('afterend', `<section class="calif-card nuevo"><b>Empieza el parcial ${n + 1}</b>
    <p class="hint">¿Cambia algo en este parcial? Ajusta lo que necesites; la lista de alumnos y todo lo capturado se conservan.</p>
    <div class="row"><button class="btn" data-c="criterios">Porcentajes y secciones</button><button class="btn" data-c="alumnos">Altas y bajas de alumnos</button></div></section>`);
}

// ---------- Eventos ----------
root().addEventListener('click', async (e) => {
  const per = e.target.closest('[data-periodo]');
  if (per) { periodo = per.dataset.periodo === 'todo' || per.dataset.periodo === 'sem' ? per.dataset.periodo : Number(per.dataset.periodo); filtro = 'all'; renderCalif(); return; }
  const f = e.target.closest('[data-f]');
  if (f) { filtro = f.dataset.f; renderCalif(); return; }
  const li = e.target.closest('[data-row]');
  if (li) {
    const { openStudent, renderStudent } = await import('./student-report.js');
    openStudent();
    ui.studentRow = Number(li.dataset.row);
    renderStudent();
    return;
  }
  const b = e.target.closest('[data-c]');
  if (!b) return;
  const c = b.dataset.c;
  if (c === 'exportar') openExportar();
  if (c === 'previa') openPreview(typeof periodo === 'number' ? periodo : 'todo');
  if (c === 'buscar') { const { openStudent } = await import('./student-report.js'); openStudent(); }
  if (c === 'cerrar') cerrarParcial(Number(b.dataset.n));
  if (c === 'reabrir') {
    const n = Number(b.dataset.n);
    const ci = cierreActivo(group(), 'parcial', n);
    if (ci && (await confirmBulk(`¿Reabrir el parcial ${n}?`, 'Sus calificaciones vuelven a calcularse con lo capturado y los porcentajes actuales. Queda registrado que se reabrió.', 'Sí, reabrir', true))) {
      reabrir(ci);
      toast(`Parcial ${n} reabierto`);
      renderCalif();
    }
  }
  if (c === 'criterios' || c === 'fechas') abrir(() => import('./grades.js'), 'openGrades');
  if (c === 'semestral') abrir(() => import('./semestral.js'), 'openSemestral');
  if (c === 'alumnos') irA('grupo');
});
// Al cerrar las pantallas de criterios o semestral, se actualiza lo que se ve.
['grades', 'semestral', 'student'].forEach((id) => { const d = $(`#${id}`); if (d) d.addEventListener('close', () => { if (ui.vista === 'calif') renderCalif(); }); });
