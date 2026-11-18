import { exportGroupReport, groupReport, round1 } from '../lib/excel-core.js';
import { gname, group, ui } from './state.js';
import { openStudent, renderStudent } from './student-report.js';
import { $, deliverFile, escapeHtml, loadExcel, toast } from './util.js';

// ---------- Resumen del grupo (semáforo) ----------
export let grFilter = 'all';
export const ST_LABEL = { rojo: 'En riesgo', amarillo: 'Observación', verde: 'Bien', none: 'Sin datos' };
export function openGroupReport() {
  grFilter = 'all';
  renderGroupReport();
  $('#group-report').showModal();
  $('#gr-list').scrollTop = 0;
}
export function renderGroupReport() {
  const g = group();
  const r = groupReport(g, g.grading);
  const st = r.stats;
  $('#gr-sub').textContent = `${gname(g)}${r.graded ? ` · calificación sobre ${r.escala}` : ''}`;
  const tile = (v, label, cls = '') => `<div class="gr-tile ${cls}"><b>${v}</b><span>${label}</span></div>`;
  $('#gr-tiles').innerHTML =
    tile(st.alumnos, 'Alumnos') +
    tile(st.asistencia == null ? '—' : `${round1(st.asistencia)} %`, 'Asistencia promedio') +
    tile(st.promedio == null ? '—' : st.promedio, r.graded ? 'Promedio general' : 'Promedio (falta definir %)') +
    tile(st.rojo, 'Alumnos en riesgo', st.rojo ? 'rojo' : '') +
    tile(st.faltas, 'Faltas en total');
  const chip = (k, n, label) => `<button class="gr-chip ${k} ${grFilter === k ? 'on' : ''}" data-f="${k}">${label} <b>${n}</b></button>`;
  $('#gr-chips').innerHTML = chip('all', st.alumnos, 'Todos') + chip('rojo', st.rojo, 'Rojo') + chip('amarillo', st.amarillo, 'Amarillo') + chip('verde', st.verde, 'Verde') +
    (st.none ? chip('none', st.none, 'Sin datos') : '');
  $('#gr-criteria').innerHTML = r.criteria.map((t) => `<li>${escapeHtml(t)}</li>`).join('');
  // Primero los que más atención necesitan.
  const order = { rojo: 0, amarillo: 1, verde: 2, none: 3 };
  const rows = r.rows.filter((x) => grFilter === 'all' || x.status === grFilter)
    .sort((a, b) => order[a.status] - order[b.status] || (a.final ?? 99) - (b.final ?? 99) || (a.att ?? 101) - (b.att ?? 101));
  $('#gr-list').innerHTML = rows.length
    ? rows.map((x) => `<li data-row="${x.student.row}" class="${x.status}">` +
        `<span class="gr-dot" title="${ST_LABEL[x.status]}"></span>` +
        `<span class="gr-name"><b>${escapeHtml(x.student.name)}</b><small>${x.reasons.length ? escapeHtml(x.reasons.join(' · ')) : ST_LABEL[x.status]}</small></span>` +
        `<span class="gr-num"><b>${x.final ?? '—'}</b><small>calif.</small></span>` +
        `<span class="gr-num"><b>${x.att == null ? '—' : `${Math.round(x.att)}%`}</b><small>asist.</small></span>` +
        `<span class="gr-num"><b>${x.counts.F}</b><small>faltas</small></span></li>`).join('')
    : '<li class="empty">No hay alumnos en este color.</li>';
}
$('#gr-chips').addEventListener('click', (e) => {
  const b = e.target.closest('[data-f]');
  if (!b) return;
  grFilter = b.dataset.f;
  renderGroupReport();
});
// Tocar un alumno abre su resumen individual.
$('#gr-list').addEventListener('click', (e) => {
  const li = e.target.closest('li[data-row]');
  if (!li) return;
  $('#group-report').close();
  openStudent();
  ui.studentRow = Number(li.dataset.row);
  renderStudent();
});
// Vista previa: genera el mismo Excel que se exporta y lo dibuja celda por celda (colores, negritas, celdas unidas).
$('#gr-preview').addEventListener('click', async () => {
  const g = group();
  const btn = $('#gr-preview');
  btn.disabled = true;
  try {
    const buf = await exportGroupReport(await loadExcel(), g, g.grading);
    const wb = new (await loadExcel()).Workbook();
    await wb.xlsx.load(buf);
    $('#gr-sheet-table').innerHTML = sheetToHtml(wb.worksheets[0]);
    $('#gr-sheet-sub').textContent = `Resumen ${gname(g)}.xlsx`;
    $('#gr-sheet').showModal();
    $('#gr-sheet .preview-scroll').scrollTop = 0;
  } catch (err) {
    console.error(err);
    toast('No se pudo generar la vista previa');
  } finally {
    btn.disabled = false;
  }
});
$('#gr-sheet-export').addEventListener('click', () => $('#gr-export').click());
export function sheetToHtml(ws) {
  const nCols = ws.columnCount;
  const nRows = ws.rowCount;
  const colName = (c) => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  // Celdas unidas: la primera lleva colspan/rowspan y las demás no se dibujan.
  const span = {};
  const skip = new Set();
  (ws.model.merges || []).forEach((rg) => {
    const [a, b] = rg.split(':').map((ref) => ws.getCell(ref));
    span[`${a.row},${a.col}`] = { cs: b.col - a.col + 1, rs: b.row - a.row + 1 };
    for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) if (r !== a.row || c !== a.col) skip.add(`${r},${c}`);
  });
  const argb = (x) => (x && x.argb ? `#${x.argb.slice(-6)}` : null);
  const text = (v) => {
    if (v == null) return '';
    if (typeof v === 'object' && v.richText) return v.richText.map((t) => t.text).join('');
    return String(v);
  };
  let html = '<thead><tr><th class="rownum"></th>';
  for (let c = 1; c <= nCols; c++) {
    // Ancho proporcional al de Excel, más compacto para que en celular se vean los números junto a los nombres.
    const w = Math.min(Math.round((ws.getColumn(c).width || 10) * 5.5), 190);
    html += `<th style="min-width:${w}px;max-width:${w}px">${colName(c)}</th>`;
  }
  html += '</tr></thead><tbody>';
  for (let r = 1; r <= nRows; r++) {
    html += `<tr><td class="rownum">${r}</td>`;
    for (let c = 1; c <= nCols; c++) {
      const key = `${r},${c}`;
      if (skip.has(key)) continue;
      const cell = ws.getCell(r, c);
      const st = [];
      const bg = cell.fill && cell.fill.type === 'pattern' ? argb(cell.fill.fgColor) : null;
      if (bg) st.push(`background:${bg}`);
      const f = cell.font || {};
      if (f.bold) st.push('font-weight:700');
      if (f.color && argb(f.color)) st.push(`color:${argb(f.color)}`);
      if (f.size && f.size > 12) st.push(`font-size:${f.size}px`);
      const al = (cell.alignment || {}).horizontal;
      if (al === 'center') st.push('text-align:center');
      if (cell.border && cell.border.top) st.push('border:1px solid #999');
      const sp = span[key];
      html += `<td${sp ? ` colspan="${sp.cs}" rowspan="${sp.rs}"` : ''}${st.length ? ` style="${st.join(';')}"` : ''}>${escapeHtml(text(cell.value)).replace(/\n/g, '<br>')}</td>`;
    }
    html += '</tr>';
  }
  return html + '</tbody>';
}
$('#gr-export').addEventListener('click', async () => {
  const g = group();
  const btn = $('#gr-export');
  btn.disabled = true;
  btn.textContent = 'Generando…';
  try {
    const out = await exportGroupReport(await loadExcel(), g, g.grading);
    const safe = (t) => String(t).replace(/[\\/:*?"<>|]+/g, ' ').trim();
    await deliverFile(out, `Resumen ${safe(gname(g))}.xlsx`);
    toast('Resumen del grupo generado');
  } catch (err) {
    console.error(err);
    toast('Error al generar el Excel');
  } finally {
    btn.disabled = false;
    btn.textContent = 'Exportar Excel';
  }
});
