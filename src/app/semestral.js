// ---------- Calificación semestral ----------
// Cada maestro elige cómo sale (por grupo): promedio de los parciales, o parciales + examen final
// (con exentos opcionales). Al cerrar el semestre las calificaciones quedan definitivas.
import { DEFAULT_GRADING, SEMESTRAL, periodosListos, round1, semestral } from '../lib/excel-core.js';
import { confirmBulk } from './dialogs.js';
import { cerrar, cierreActivo, fechaCorta, reabrir } from './cierres.js';
import { gname, group, save, state } from './state.js';
import { $, escapeHtml, toast } from './util.js';
import { openGrades } from './grades.js';

const cfgDe = (g) => (g.semestral = Object.assign({}, SEMESTRAL, g.semestral || {}, { final: (g.semestral && g.semestral.final) || {} }));
const fmt = (v) => (v == null ? '—' : round1(v));

export function openSemestral() {
  render();
  $('#semestral').showModal();
}

function render() {
  const g = group();
  const gr = Object.assign({}, DEFAULT_GRADING, g.grading);
  const escala = Number(gr.escala) || 10;
  const per = state.periodos;
  $('#sem-sub').textContent = `${gname(g)} · sobre ${escala}${per && per.semestre ? ` · ${per.semestre}` : ''}`;
  const body = $('#sem-body');
  if (!periodosListos(per)) {
    body.innerHTML = `<div class="cuenta-card"><b>Faltan las fechas de los parciales</b>
      <small>Defínelas en Secciones y calificaciones → Parciales del semestre.</small></div>
      <button class="btn" data-sem="ir-parciales">Definir parciales</button>`;
    $('#sem-cerrar').hidden = true;
    return;
  }
  const c = cfgDe(g);
  const cierre = cierreActivo(g, 'semestre');
  const bloqueado = !!cierre;
  const n = per.parciales.length;
  const minimo = c.minimo != null ? c.minimo : escala * 0.6;
  const opciones = `<form class="sem-opts" id="sem-opts" onsubmit="return false">
      <fieldset ${bloqueado ? 'disabled' : ''}>
        <label class="sem-radio"><input type="radio" name="modo" value="promedio" ${c.modo === 'promedio' ? 'checked' : ''}> Promedio de los ${n} parciales</label>
        <label class="sem-radio"><input type="radio" name="modo" value="final" ${c.modo === 'final' ? 'checked' : ''}> Parciales + examen final</label>
        ${c.modo === 'final' ? `<div class="sem-final">
          <label>Parciales <input name="pesoParciales" inputmode="decimal" value="${c.pesoParciales}"> %</label>
          <label>Examen final <input name="pesoFinal" inputmode="decimal" value="${c.pesoFinal}"> %</label>
          <label class="check"><input type="checkbox" name="exentos" ${c.exentoDesde != null ? 'checked' : ''}> Exentos desde
            <input name="exentoDesde" inputmode="decimal" value="${c.exentoDesde ?? (escala === 10 ? 8 : 80)}" ${c.exentoDesde == null ? 'disabled' : ''}> de promedio</label>
        </div>` : ''}
        <div class="sem-final">
          <label>Aprueba con <input name="minimo" inputmode="decimal" value="${minimo}"></label>
          <label>Reportar <select name="redondeo"><option value="decimal" ${c.redondeo !== 'entero' ? 'selected' : ''}>con un decimal</option><option value="entero" ${c.redondeo === 'entero' ? 'selected' : ''}>en entero</option></select></label>
          ${escala === 10 ? `<label class="check"><input type="checkbox" name="reprobado5" ${c.reprobado5 ? 'checked' : ''}> Reprobados como 5</label>` : ''}
        </div>
      </fieldset>
    </form>`;
  const conFinal = c.modo === 'final';
  let tabla = `<thead><tr><th class="stick">No.</th><th class="stick name">Nombre</th><th class="final">Semestral</th>
    ${per.parciales.map((p) => `<th>P${p.numero}</th>`).join('')}<th>Prom.</th>${conFinal ? '<th>Final<br><small>0–100</small></th>' : ''}<th></th></tr></thead><tbody>`;
  let aprobados = 0;
  let reprobados = 0;
  let faltan = 0;
  for (const s of g.students) {
    const r = cierre && cierre.calificaciones[s.row] ? cierre.calificaciones[s.row] : semestral(g, s.row, gr, per, c);
    if (r.aprobado === true) aprobados++;
    if (r.aprobado === false) reprobados++;
    if (r.faltaFinal) faltan++;
    const estado = r.faltaFinal ? '<span class="sem-falta">falta final</span>' : r.exento ? '<span class="sem-ok">exento</span>'
      : r.aprobado === true ? '<span class="sem-ok">✓</span>' : r.aprobado === false ? '<span class="sem-mal">✗</span>' : '';
    tabla += `<tr><td class="stick">${escapeHtml(s.no)}</td><td class="stick name">${escapeHtml(s.name)}</td>
      <td class="v final ${r.aprobado === false ? 'sem-rep' : ''}">${fmt(r.reportada)}</td>
      ${(r.parciales || []).map((v) => `<td class="v">${fmt(v)}</td>`).join('')}<td class="v">${fmt(r.promedio)}</td>
      ${conFinal ? `<td class="v">${r.exento ? '—' : `<input class="exam-in sem-in" data-row="${s.row}" inputmode="decimal" maxlength="5" value="${c.final[s.row] ?? ''}" ${bloqueado ? 'disabled' : ''} aria-label="Examen final de ${escapeHtml(s.name)}">`}</td>` : ''}
      <td>${estado}</td></tr>`;
  }
  tabla += '</tbody>';
  const resumen = `<p class="sem-resumen">Aprobados: <b>${aprobados}</b> · Reprobados: <b>${reprobados}</b>${conFinal ? ` · Falta examen final: <b>${faltan}</b>` : ''}</p>`;
  const aviso = cierre ? `<div class="per-cierre"><span>🔒 Semestre cerrado el ${fechaCorta(cierre.cerrado_en)}: calificaciones definitivas.</span><button class="link-btn" data-sem="reabrir">Reabrir</button></div>` : '';
  body.innerHTML = aviso + opciones + resumen + `<div class="preview-scroll"><table class="sheet grades-table" id="sem-table">${tabla}</table></div>`;
  const t = $('#sem-table');
  const w = t.querySelector('thead th.stick').offsetWidth;
  t.querySelectorAll('tr').forEach((tr) => { tr.children[0].style.left = '0px'; tr.children[1].style.left = `${w}px`; });
  const btn = $('#sem-cerrar');
  btn.hidden = bloqueado;
  btn.textContent = 'Cerrar semestre';
}

$('#sem-body').addEventListener('change', (e) => {
  const g = group();
  const c = cfgDe(g);
  const el = e.target;
  if (el.classList.contains('sem-in')) {
    const t = el.value.trim().replace(',', '.');
    if (t === '') delete c.final[el.dataset.row];
    else {
      const v = Number(t);
      if (!Number.isFinite(v) || v < 0 || v > 100) { toast('El examen final va de 0 a 100'); el.value = c.final[el.dataset.row] ?? ''; return; }
      c.final[el.dataset.row] = v;
    }
  } else if (el.form && el.form.id === 'sem-opts') {
    const f = el.form;
    c.modo = f.elements.modo.value;
    if (f.elements.pesoParciales) {
      c.pesoParciales = Number(f.elements.pesoParciales.value) || 0;
      c.pesoFinal = Number(f.elements.pesoFinal.value) || 0;
      c.exentoDesde = f.elements.exentos.checked ? Number(f.elements.exentoDesde.value) || null : null;
    }
    c.minimo = f.elements.minimo.value === '' ? null : Number(f.elements.minimo.value);
    c.redondeo = f.elements.redondeo.value;
    c.reprobado5 = !!(f.elements.reprobado5 && f.elements.reprobado5.checked);
  } else return;
  save();
  const foco = el.classList.contains('sem-in') ? el.dataset.row : null;
  render();
  if (foco) {
    const all = [...document.querySelectorAll('#sem-table .sem-in')];
    const i = all.findIndex((x) => x.dataset.row === foco);
    if (all[i + 1]) all[i + 1].focus();
  }
});
$('#sem-body').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.classList.contains('sem-in')) { e.preventDefault(); e.target.blur(); }
});
$('#sem-body').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-sem]');
  if (!b) return;
  if (b.dataset.sem === 'ir-parciales') { $('#semestral').close(); openGrades(); $('#periodos').open = true; return; }
  if (b.dataset.sem === 'reabrir') {
    const c = cierreActivo(group(), 'semestre');
    if (!c || !(await confirmBulk('¿Reabrir el semestre?', 'Las calificaciones vuelven a calcularse. Queda registrado que se reabrió.', 'Sí, reabrir', true))) return;
    reabrir(c);
    render();
  }
});
$('#sem-cerrar').addEventListener('click', async () => {
  const g = group();
  const gr = Object.assign({}, DEFAULT_GRADING, g.grading);
  const c = cfgDe(g);
  const filas = g.students.map((s) => [s, semestral(g, s.row, gr, state.periodos, c)]);
  const faltan = filas.filter(([, r]) => r.faltaFinal).length;
  if (!(await confirmBulk('¿Cerrar el semestre?', `${faltan ? `${faltan} alumno(s) aún no tienen examen final. ` : ''}Las calificaciones semestrales de ${gname(g)} quedan definitivas. Se puede reabrir si hay que corregir.`, 'Sí, cerrar semestre'))) return;
  const cal = Object.fromEntries(filas.map(([s, r]) => [s.row, r]));
  cerrar(g, 'semestre', null, cal);
  toast('Semestre cerrado');
  render();
});
