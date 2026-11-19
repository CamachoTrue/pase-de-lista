import { exportStudent, finalGrade, round1, studentReport, weightOf } from '../lib/excel-core.js';
import { gname, group, ui } from './state.js';
import { $, deliverFile, escapeHtml, loadExcel, normalize, toast } from './util.js';
import { alElegirPeriodo, botonesPeriodo, califCerrada, vistaPeriodo } from './periodo-vista.js';

// ---------- Resumen por alumno ----------
// Para quejas o dudas del alumno o de los padres: todo lo capturado de un alumno y su Excel individual.
export function openStudent() {
  ui.studentRow = null;
  $('#student-search').value = '';
  renderStudent();
  $('#student').showModal();
}
export function renderStudent() {
  const g0 = group();
  const v = vistaPeriodo(g0);
  const g = v.g;
  const picked = g.students.find((s) => s.row === ui.studentRow);
  $('#student-periodos').innerHTML = botonesPeriodo(g0);
  $('#student-pick').hidden = !!picked;
  $('#student-report').hidden = !picked;
  $('#student-back').hidden = !picked;
  $('#student-export').hidden = !picked;
  $('#student-close').hidden = !!picked;
  if (!picked) {
    $('#student-sub').textContent = `${gname(g)}${v.etiqueta ? ` · ${v.etiqueta}` : ''} · elige un alumno`;
    const q = normalize($('#student-search').value);
    const shown = g.students.filter((s) => !q || normalize(s.name).includes(q));
    $('#student-list').innerHTML = shown.length
      ? shown.map((s) => {
          const c = califCerrada(v, s.row);
          const f = c !== undefined ? c : finalGrade(g, s.row, v.grading).final;
          return `<li data-row="${s.row}"><span class="no">${escapeHtml(s.no)}</span><span class="nm">${escapeHtml(s.name)}</span>${f != null ? `<span class="cal">${f}</span>` : ''}</li>`;
        }).join('')
      : '<li class="empty">Ningún alumno coincide con la búsqueda.</li>';
    return;
  }
  const r = studentReport(g, ui.studentRow, v.grading);
  const cerrada = califCerrada(v, picked.row);
  if (cerrada !== undefined) r.final = cerrada;
  $('#student-sub').textContent = `${gname(g)}${v.etiqueta ? ` · ${v.etiqueta}` : ''}`;
  const pct = (v) => (v == null ? 'Sin datos' : `${round1(v)} %`);
  $('#student-report').innerHTML =
    `<div class="sr-head"><div><b>${escapeHtml(picked.name)}</b><small>No. ${escapeHtml(picked.no)} · ${escapeHtml(gname(g))}</small></div>` +
    `<div class="sr-final">${r.final == null ? '—' : r.final}<small>${r.graded ? `sobre ${r.grading.escala}` : 'sin porcentajes'}</small></div></div>` +
    '<table class="sheet sr-parts"><thead><tr><th>Sección</th><th>Vale</th><th>Resultado</th></tr></thead><tbody>' +
    r.sections.map((p) => `<tr><td><b>${escapeHtml(p.name)}</b><small>${escapeHtml(r.detail[p.id])}</small></td>` +
      `<td class="v">${weightOf(r.grading, p) ? `${weightOf(r.grading, p)} %` : '—'}</td><td class="v">${pct(r.parts[p.id])}</td></tr>`).join('') +
    '</tbody></table>' +
    '<h3 class="sr-title">Detalle por día</h3>' +
    (r.days.length
      ? '<div class="preview-scroll"><table class="sheet sr-days"><thead><tr><th>Fecha</th>' +
        r.sections.map((p) => `<th class="wrap">${escapeHtml(p.name)}</th>`).join('') + '</tr></thead><tbody>' +
        r.days.map((d) => `<tr><td class="nowrap">${d.label.replace(/\/\d\d(\d\d)$/, '/$1')}</td>` +
          r.sections.map((p) => {
            const c = d.cells[p.id];
            if (!c) return '<td></td>';
            if (p.type === 'trab') return `<td>${c.list.map((w) => `<span class="${Number(w.v) === 0 ? 'F' : ''}">${escapeHtml(w.title)}: ${w.v}</span>`).join('<br>')}</td>`;
            return `<td class="v ${c.tone || ''}">${escapeHtml(String(c.short))}</td>`;
          }).join('') + '</tr>').join('') +
        '</tbody></table></div>'
      : `<p class="hint">${v.etiqueta && v.sufijo ? 'No hay datos capturados de este alumno en este parcial.' : 'Aún no hay datos capturados de este alumno.'}</p>`);
}
alElegirPeriodo($('#student-periodos'), renderStudent);
$('#student-search').addEventListener('input', renderStudent);
$('#student-list').addEventListener('click', (e) => {
  const li = e.target.closest('li[data-row]');
  if (!li) return;
  ui.studentRow = Number(li.dataset.row);
  renderStudent();
  $('#student-report').scrollTop = 0;
});
$('#student-back').addEventListener('click', () => { ui.studentRow = null; renderStudent(); });
$('#student-export').addEventListener('click', async () => {
  const v = vistaPeriodo(group());
  const g = v.g;
  const s = g.students.find((x) => x.row === ui.studentRow);
  if (!s) return;
  const btn = $('#student-export');
  btn.disabled = true;
  btn.textContent = 'Generando…';
  try {
    const out = await exportStudent(await loadExcel(), g, s.row, v.grading);
    const safe = (t) => String(t).replace(/[\\/:*?"<>|]+/g, ' ').trim();
    await deliverFile(out, `${safe(s.name)} - ${safe(gname(g))}${v.sufijo || ''} - resumen.xlsx`);
    toast('Excel del alumno generado');
  } catch (err) {
    console.error(err);
    toast('Error al generar el Excel');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Exportar Excel';
  }
});
