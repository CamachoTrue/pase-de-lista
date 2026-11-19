import { ASIST_ID, TYPES, enPeriodo, finalGrade, newSectionId, normalizeGroup, periodosListos, round1, secMin, sectionsOf, weightOf } from '../lib/excel-core.js';
import { cerrar, cierreActivo, fechaCorta, reabrir } from './cierres.js';
import { confirmBulk } from './dialogs.js';
import { refresh } from './main-view.js';
import { gname, group, save, secById, state } from './state.js';
import { $, escapeHtml, normalize, toast } from './util.js';
import { paqueteGuardado, rolGuardado } from '../lib/nube.js';

// ---------- Secciones y calificaciones ----------
// Cada grupo tiene sus secciones (nombre + plantilla) y cuánto vale cada una. Todo se edita en un
// borrador: la tabla se actualiza al instante, pero solo se guarda con "Guardar".
export let gradeDraft = null; // { sections, grading }
export const snapshot = (g) => JSON.parse(JSON.stringify({ sections: sectionsOf(g), grading: g.grading, periodos: state.periodos || null }));
// Periodo que se ve en la tabla: 0 = todo; 1-3 = parcial.
let periodoVista = 0;
const periodosDeEscuela = () => !!(state.periodos && state.periodos.origen === 'escuela');
// Maestro de una escuela (con director): las fechas las pone solo el director. El maestro
// independiente (o sin cuenta) las pone él mismo.
const deUnaEscuela = () => paqueteGuardado() === 'empresarial';
export const sameDraft = () => JSON.stringify(gradeDraft) === JSON.stringify(snapshot(group()));
// Grupo "de prueba" con el borrador, para calcular la tabla sin tocar el grupo real.
export const draftGroup = () => Object.assign({}, group(), { sections: gradeDraft.sections });
export const typeLabel = (t) => TYPES[t].label;
export const draftTotal = () => round1(gradeDraft.sections.reduce((n, x) => n + weightOf(gradeDraft.grading, x), 0));
// Cuántos datos tiene una sección (para avisar antes de borrarla).
export function secDataCount(g, sec) {
  const d = g[sec.id] || {};
  return Object.keys(d).filter((k) => (Array.isArray(d[k]) ? d[k].some((w) => Object.keys(w.vals || {}).length) : Object.keys(d[k]).length)).length;
}

export function openGrades() {
  gradeDraft = snapshot(group());
  $('#grades-scope').textContent = `Secciones de ${gname(group())} y cuánto vale cada una. Cada grupo tiene las suyas.`;
  const f = $('#grades-form');
  f.elements.retardo.value = String(gradeDraft.grading.retardo);
  f.elements.escala.value = String(gradeDraft.grading.escala);
  pintarPeriodos();
  renderSecList();
  renderGrades();
  $('#grades').showModal();
}
// Fechas de los parciales en el formulario (deshabilitadas si las define la escuela).
function pintarPeriodos() {
  const f = $('#grades-form');
  const p = gradeDraft.periodos || {};
  const pars = p.parciales || [];
  f.elements.semestre.value = p.semestre || '';
  [1, 2, 3].forEach((n) => {
    const x = pars.find((y) => y.numero === n) || {};
    f.elements[`p${n}i`].value = x.inicio || '';
    f.elements[`p${n}f`].value = x.fin || '';
  });
  const deEscuela = periodosDeEscuela();
  const bloqueado = deEscuela || deUnaEscuela();
  ['semestre', 'p1i', 'p1f', 'p2i', 'p2f', 'p3i', 'p3f'].forEach((k) => (f.elements[k].disabled = bloqueado));
  $('#per-nota').textContent = bloqueado && rolGuardado() === 'director'
    ? 'Como director, las fechas de los parciales las pones en Mi escuela → Parciales (valen para toda la escuela).'
    : deEscuela
    ? 'Las fechas de los parciales las puso tu director para toda la escuela. Si alguna está mal, pídele que la cambie.'
    : bloqueado ? 'En tu escuela, las fechas de los parciales las pone el director. Aún no las ha puesto; en cuanto lo haga aparecerán aquí solas.'
    : 'Las capturas se acomodan solas en su parcial según la fecha. Valen para todos tus grupos.';
  $('#per-resumen').textContent = periodosListos(p) ? `· ${p.semestre || 'definidos'}` : '· sin definir';
}
function leerPeriodos() {
  const f = $('#grades-form');
  const parciales = [1, 2, 3].map((n) => ({ numero: n, inicio: f.elements[`p${n}i`].value, fin: f.elements[`p${n}f`].value }))
    .filter((x) => x.inicio || x.fin);
  gradeDraft.periodos = parciales.length || f.elements.semestre.value.trim()
    ? { semestre: f.elements.semestre.value.trim(), parciales, origen: 'maestro' } : null;
}
$('#grades-form').addEventListener('input', (e) => {
  if (!/^(semestre|p\d[if])$/.test(e.target.name || '')) return;
  leerPeriodos();
  renderGrades();
});
$('#per-tabs').addEventListener('click', (e) => {
  const b = e.target.closest('[data-p]');
  if (!b) return;
  periodoVista = Number(b.dataset.p);
  renderGrades();
});
$('#per-cierre').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-cierre]');
  if (!b) return;
  const g = group();
  if (!sameDraft()) { toast('Guarda los cambios antes de cerrar el parcial'); return; }
  if (b.dataset.cierre === 'cerrar') {
    const p = state.periodos.parciales.find((x) => x.numero === periodoVista);
    if (!(await confirmBulk(`¿Cerrar el parcial ${periodoVista}?`, 'Las calificaciones de este parcial quedan guardadas como definitivas para este grupo. Si hay que corregir algo, se puede reabrir.', 'Sí, cerrar'))) return;
    const vista = enPeriodo(g, p.inicio, p.fin);
    const cal = {};
    g.students.forEach((s) => { const r = finalGrade(vista, s.row, g.grading); cal[s.row] = { final: r.final, partes: r.parts }; });
    cerrar(g, 'parcial', periodoVista, cal, g.grading);
    toast(`Parcial ${periodoVista} cerrado`);
  } else {
    const c = cierreActivo(g, 'parcial', periodoVista);
    if (!c || !(await confirmBulk(`¿Reabrir el parcial ${periodoVista}?`, 'Las calificaciones vuelven a calcularse con lo capturado. Queda registrado que se reabrió.', 'Sí, reabrir', true))) return;
    reabrir(c);
    toast(`Parcial ${periodoVista} reabierto`);
  }
  renderGrades();
});
export function renderSecList() {
  const g = group();
  $('#sec-list').innerHTML = gradeDraft.sections.map((x, i) => {
    const isNew = !sectionsOf(g).some((o) => o.id === x.id);
    const w = gradeDraft.grading.w[x.id];
    const typeBox = isNew
      ? `<select data-f="type" aria-label="Plantilla">${Object.keys(TYPES).filter((t) => t !== 'asist').map((t) => `<option value="${t}" ${t === x.type ? 'selected' : ''}>${escapeHtml(typeLabel(t))}</option>`).join('')}</select>`
      : `<span class="sec-type">${escapeHtml(typeLabel(x.type))}</span>`;
    return `<div class="sec-row ${isNew ? 'new' : ''}" data-i="${i}">
      <div class="sec-top">
        <input class="sec-name" data-f="name" maxlength="30" value="${escapeHtml(x.name)}" placeholder="Nombre de la sección" aria-label="Nombre de la sección">
        <span class="sec-w"><input data-f="w" inputmode="decimal" maxlength="5" value="${w || ''}" aria-label="Porcentaje de ${escapeHtml(x.name)}">%</span>
        ${x.type === 'asist' ? '<span class="sec-del-ph"></span>' : `<button type="button" class="icon-btn sec-del" data-del="${i}" aria-label="Quitar ${escapeHtml(x.name)}">✕</button>`}
      </div>
      <div class="sec-sub">${typeBox}${x.type === 'sino' ? `<label class="sec-min">100 % con <input data-f="min" class="mini-num" inputmode="numeric" maxlength="3" value="${secMin(x)}"> o más “Sí”</label>` : ''}</div>
    </div>`;
  }).join('');
}
$('#sec-list').addEventListener('input', (e) => {
  const el = e.target.closest('[data-f]');
  const row = e.target.closest('.sec-row');
  if (!el || !row) return;
  const x = gradeDraft.sections[Number(row.dataset.i)];
  const f = el.dataset.f;
  if (f === 'name') x.name = el.value.replace(/[:\n]/g, ' ').replace(/\s+/g, ' ').trimStart();
  if (f === 'w') {
    const n = Number(el.value.replace(',', '.'));
    if (Number.isFinite(n) && n > 0) gradeDraft.grading.w[x.id] = n; else delete gradeDraft.grading.w[x.id];
  }
  if (f === 'min') { const n = Math.round(Number(el.value)); if (Number.isFinite(n) && n >= 1) x.min = n; }
  if (f === 'type') {
    x.type = el.value;
    if (x.type === 'sino') x.min = x.min || 4; else delete x.min;
    renderSecList();
  }
  renderGrades();
});
$('#sec-list').addEventListener('click', (e) => {
  const b = e.target.closest('[data-del]');
  if (!b) return;
  const x = gradeDraft.sections[Number(b.dataset.del)];
  gradeDraft.sections.splice(Number(b.dataset.del), 1);
  delete gradeDraft.grading.w[x.id];
  renderSecList();
  renderGrades();
  const n = secDataCount(group(), x);
  toast(n ? `“${x.name}” se quitará al guardar (tiene datos de ${n} día(s))` : `“${x.name || 'Sección'}” quitada`);
});
$('#sec-add').addEventListener('click', () => {
  gradeDraft.sections.push({ id: newSectionId(), name: '', type: 'num' });
  renderSecList();
  renderGrades();
  const rows = document.querySelectorAll('#sec-list .sec-row');
  const last = rows[rows.length - 1];
  last.scrollIntoView({ block: 'nearest' });
  last.querySelector('.sec-name').focus();
});
$('#grades-form').addEventListener('change', (e) => {
  if (e.target.name === 'retardo') gradeDraft.grading.retardo = Number(e.target.value);
  if (e.target.name === 'escala') gradeDraft.grading.escala = Number(e.target.value);
  if (e.target.name === 'retardo' || e.target.name === 'escala') renderGrades();
});
$('#grades-save').addEventListener('click', async () => {
  const g = group();
  const list = gradeDraft.sections;
  list.forEach((x) => (x.name = x.name.trim()));
  const empty = list.find((x) => !x.name);
  if (empty) { toast('Ponle nombre a todas las secciones'); return; }
  const seen = new Set();
  const dup = list.find((x) => { const k = normalize(x.name); if (seen.has(k)) return true; seen.add(k); return false; });
  if (dup) { toast(`Hay dos secciones llamadas “${dup.name}”`); return; }
  const total = draftTotal();
  if (total !== 100 && total !== 0) { toast(`Los porcentajes suman ${total} %; deben sumar 100 %`); return; }
  // Secciones quitadas que tenían datos: se avisa antes de borrarlas.
  const gone = sectionsOf(g).filter((o) => !list.some((x) => x.id === o.id));
  const withData = gone.map((o) => [o, secDataCount(g, o)]).filter(([, n]) => n);
  if (withData.length && !(await confirmBulk(
    `¿Quitar ${withData.map(([o]) => `“${o.name}”`).join(', ')}?`,
    `Se borrarán sus datos capturados (${withData.map(([o, n]) => `${o.name}: ${n} día(s)`).join(', ')}). No se puede deshacer.`,
    'Sí, quitar', true))) return;
  gone.forEach((o) => delete g[o.id]);
  g.sections = JSON.parse(JSON.stringify(list));
  g.grading = JSON.parse(JSON.stringify(gradeDraft.grading));
  Object.keys(g.grading.w).forEach((id) => { if (!g.sections.some((x) => x.id === id)) delete g.grading.w[id]; });
  normalizeGroup(g);
  if (!periodosDeEscuela() && !deUnaEscuela()) state.periodos = gradeDraft.periodos ? JSON.parse(JSON.stringify(gradeDraft.periodos)) : null;
  gradeDraft = snapshot(g);
  if (!secById(state.ui.tab)) state.ui.tab = ASIST_ID;
  save();
  renderSecList();
  renderGrades();
  refresh();
  toast(`Secciones y porcentajes de ${gname(g)} guardados`);
});
// Al cerrar sin guardar se descartan los cambios (avisando).
$('#grades').addEventListener('close', () => {
  if (gradeDraft && !sameDraft()) toast('No se guardaron los cambios');
  gradeDraft = null;
});
export function renderGrades() {
  const g = group();
  const dg = draftGroup();
  const gr = gradeDraft.grading;
  const total = draftTotal();
  const saved = sameDraft();
  const tot = $('#weights-total');
  const okTotal = total === 100 || total === 0;
  tot.textContent = total === 0
    ? (saved ? 'Sin porcentajes · Guardado ✓' : 'Sin porcentajes (no se calcula calificación) · Falta guardar')
    : total !== 100
      ? `Total: ${total} % — debe sumar 100 %`
      : saved ? 'Total: 100 % · Guardado ✓' : 'Total: 100 % · Falta guardar';
  tot.className = 'weights-total ' + (!okTotal ? 'bad' : saved ? 'ok' : 'warn');
  const btn = $('#grades-save');
  btn.disabled = saved;
  btn.textContent = saved ? 'Guardado ✓' : 'Guardar';
  $('#grades-sub').textContent = `${gname(g)} · ${g.students.length} alumnos · sobre ${gr.escala}`;
  // Pestañas de periodo (si hay parciales definidos) y su cierre.
  const per = gradeDraft.periodos;
  const listos = periodosListos(per);
  if (!listos) periodoVista = 0;
  $('#per-tabs').innerHTML = listos
    ? [0, ...per.parciales.map((x) => x.numero)].map((n) => `<button type="button" class="per-tab ${n === periodoVista ? 'on' : ''}" data-p="${n}">${n ? `Parcial ${n}${cierreActivo(g, 'parcial', n) ? ' 🔒' : ''}` : 'Todo'}</button>`).join('')
    : '';
  const cierre = periodoVista ? cierreActivo(g, 'parcial', periodoVista) : null;
  $('#per-cierre').innerHTML = !periodoVista ? ''
    : cierre ? `<span>🔒 Cerrado el ${fechaCorta(cierre.cerrado_en)}: calificaciones definitivas.</span><button type="button" class="link-btn" data-cierre="reabrir">Reabrir</button>`
      : `<button type="button" class="btn per-cerrar" data-cierre="cerrar">Cerrar parcial ${periodoVista}</button>`;
  const pv = periodoVista && listos ? per.parciales.find((x) => x.numero === periodoVista) : null;
  const vista = pv ? enPeriodo(dg, pv.inicio, pv.fin) : dg;
  const parts = dg.sections.filter((x) => weightOf(gr, x) > 0);
  if (!parts.length) {
    $('#grades-table').innerHTML = '<tbody><tr><td class="grades-empty">Escribe cuánto vale cada sección para ver las calificaciones.</td></tr></tbody>';
    return;
  }
  const fmt = (v) => (v == null ? '—' : round1(v));
  // La calificación final va junto al nombre para que se vea sin desplazar en el celular.
  let html = '<thead><tr><th class="stick">No.</th><th class="stick name">Nombre</th><th class="final">Calif.</th>' +
    parts.map((x) => `<th>${escapeHtml(x.name || 'Sin nombre')}<br><small>${weightOf(gr, x)}%${x.type === 'sino' ? ` · mín. ${secMin(x)}` : ''}</small></th>`).join('') +
    '</tr></thead><tbody>';
  for (const st of g.students) {
    const r = cierre && cierre.calificaciones[st.row] ? { final: cierre.calificaciones[st.row].final, parts: cierre.calificaciones[st.row].partes || {} } : finalGrade(vista, st.row, gr);
    html += `<tr><td class="stick">${escapeHtml(st.no)}</td><td class="stick name">${escapeHtml(st.name)}</td><td class="v final">${fmt(r.final)}</td>` +
      parts.map((x) => `<td class="v">${fmt(r.parts[x.id])}</td>`).join('') +
      '</tr>';
  }
  const t = $('#grades-table');
  t.innerHTML = html + '</tbody>';
  // No. y Nombre fijos al desplazar a la derecha.
  const w = t.querySelector('thead th.stick').offsetWidth;
  t.querySelectorAll('tr').forEach((tr) => {
    tr.children[0].style.left = '0px';
    tr.children[1].style.left = `${w}px`;
  });
}
