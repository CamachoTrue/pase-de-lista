import { READ, STATUSES, cellText, cierreVigente, enPeriodo, exportGroupReport, exportWorkbook, groupDates, periodosListos, sheetLayout, worksOf } from '../lib/excel-core.js';
import { LABELS, gname, group, optionOf, secs, state } from './state.js';
import { $, b64ToBuf, deliverFile, escapeHtml, loadExcel, toast } from './util.js';

// ---------- Exportar ----------
$('#done-export').addEventListener('click', () => openExportar());
$('#preview-export').addEventListener('click', () => doExport($('#preview-export'), alcance));

// ---------- Qué exportar: todo el semestre, un parcial o el resumen del grupo ----------
let alcance = 'todo';
const parciales = () => (periodosListos(state.periodos) ? state.periodos.parciales : []);
// Grupos tal como se exportan: todo, o solo las fechas de un parcial (con los criterios con que se cerró).
function preparar(groups, a) {
  const p = parciales().find((x) => x.numero === a);
  if (!p) return { groups, periodos: state.periodos };
  return {
    periodos: null,
    groups: groups.map((g) => {
      const v = enPeriodo(g, p.inicio, p.fin);
      const c = cierreVigente(g, 'parcial', p.numero);
      v.grading = (c && c.criterios) || g.grading;
      return v;
    }),
  };
}
const fechaCortita = (iso) => new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
export function openExportar() {
  const g = group();
  const opc = [['todo', parciales().length ? 'Todo el semestre' : 'Todo lo capturado', parciales().length ? 'Todas las fechas, con columnas P1, P2, P3 y semestral' : 'Todas las fechas y la calificación']];
  for (const p of parciales()) {
    const c = cierreVigente(g, 'parcial', p.numero);
    opc.push([p.numero, `Solo el parcial ${p.numero}`, `Del ${fechaCortita(p.inicio)} al ${fechaCortita(p.fin)}${c ? ' · cerrado' : ''}`]);
  }
  opc.push(['resumen', 'Resumen del grupo (semáforo)', `Solo ${gname(g)}: calificación, asistencia y riesgo de cada alumno`]);
  if (!opc.some(([k]) => k === alcance)) alcance = 'todo';
  $('#exp-opciones').innerHTML = opc.map(([k, t, sub]) => `<label class="exp-op"><input type="radio" name="exp" value="${k}" ${k === alcance ? 'checked' : ''}>
    <span><b>${escapeHtml(t)}</b><small>${escapeHtml(sub)}</small></span></label>`).join('');
  $('#exp-previa').hidden = alcance === 'resumen';
  $('#exportar').showModal();
}
$('#exp-opciones').addEventListener('change', (e) => {
  const v = e.target.value;
  alcance = v === 'todo' || v === 'resumen' ? v : Number(v);
  $('#exp-previa').hidden = alcance === 'resumen';
});
$('#exp-previa').addEventListener('click', () => { $('#exportar').close(); openPreview(alcance); });
$('#exp-ok').addEventListener('click', async () => {
  const btn = $('#exp-ok');
  if (alcance === 'resumen') {
    btn.disabled = true;
    try {
      const g = group();
      await deliverFile(await exportGroupReport(await loadExcel(), g, g.grading), `Resumen ${gname(g)}.xlsx`);
      toast('Excel generado');
    } catch (err) { console.error(err); toast('Error al generar el Excel'); } finally { btn.disabled = false; }
  } else await doExport(btn, alcance);
  $('#exportar').close();
});
$('#done-preview').addEventListener('click', openPreview);
// ---------- Vista previa ----------
// Se genera el mismo Excel que al exportar y se vuelve a leer, así lo que se ve es exactamente lo que se descarga.
export async function openPreview(a = alcance) {
  alcance = a;
  const prep = preparar([group()], a);
  const g = prep.groups[0];
  let ws;
  try {
    const out = await exportWorkbook(await loadExcel(), b64ToBuf(state.fileB64), [g], state.codes, null, prep.periodos);
    const wb = new (await loadExcel()).Workbook();
    await wb.xlsx.load(out);
    ws = wb.getWorksheet(g.sheet);
  } catch (err) {
    console.error(err);
    toast('No se pudo generar la vista previa');
    return;
  }

  const dates = groupDates(g);
  const layout = sheetLayout(g, g.grading, prep.periodos);
  const byCol = new Map(layout.map((x) => [x.col, x]));
  const lastCol = Math.max(g.nameCol, ...layout.map((x) => x.col));
  const isCur = (c) => byCol.has(c) && byCol.get(c).date === state.ui.date;
  const colLetter = (c) => { let s = ''; for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };
  // Se toma de lo capturado (no del número), porque A, R y J pueden tener el mismo valor.
  const colorOf = (x, r) => {
    if (x.kind === 'disc-count') return discCount(x.sec, r) > 0 ? 'F' : '';
    if (x.kind !== 'day' && x.kind !== 'trab') return '';
    const v = x.kind === 'trab' ? (worksOf(g, x.date, x.sec.id)[x.work].vals || {})[r] : ((g[x.sec.id] || {})[x.date] || {})[r];
    if (v == null) return '';
    const o = optionOf(x.sec.id, v);
    return o ? o.color : '';
  };
  const discCount = (sec, r) => Object.values(g[sec.id] || {}).filter((d) => d[r] === 0).length;
  const text = (v) => {
    if (v instanceof Date) return `${String(v.getUTCDate()).padStart(2, '0')}/${String(v.getUTCMonth() + 1).padStart(2, '0')}/${String(v.getUTCFullYear()).slice(-2)}`;
    return cellText(v).replace(/\s+/g, ' ').trim();
  };

  let html = '<thead><tr><th class="rownum"></th>';
  for (let c = 1; c <= lastCol; c++) html += `<th class="${c <= g.nameCol ? 'stick' : ''}">${colLetter(c)}</th>`;
  html += '</tr></thead><tbody>';
  const rows = [g.headerRow, ...g.students.map((s) => s.row)];
  rows.forEach((r, i) => {
    const isHdr = i === 0;
    html += `<tr class="${isHdr ? 'hdr' : ''}"><td class="rownum">${r}</td>`;
    for (let c = 1; c <= lastCol; c++) {
      const v = ws.getCell(r, c).value;
      const t = text(v);
      const cls = [];
      if (c <= g.nameCol) cls.push('stick');
      if (c === g.nameCol) cls.push('name');
      const x = byCol.get(c);
      if (x) {
        cls.push(isHdr ? 'date-h' : 'v');
        if (!isHdr) cls.push(colorOf(x, r));
        if (x.first) cls.push('day-start');
        if (x.last) cls.push('block-end');
        if (x.kind === 'day' && x.sec.type === 'sino') cls.push('lect-col');
        if (x.kind === 'grade' && x.key === 'final') cls.push('final');
        if (x.kind === 'disc-dates') cls.push('disc-dates');
        if (isCur(c)) cls.push('cur');
      }
      // El encabezado lleva fecha y apartado en dos renglones, como en el Excel.
      const shown = isHdr && x ? escapeHtml(cellText(v)).replace(/\n/g, '<br>') : escapeHtml(t);
      html += `<td class="${cls.join(' ')}">${shown}</td>`;
    }
    html += '</tr>';
  });
  html += '</tbody>';
  $('#preview-table').innerHTML = html;

  $('#preview-title').textContent = `Vista previa · ${gname(g)}${typeof a === 'number' ? ` · parcial ${a}` : ''}`;
  $('#preview-sub').textContent = layout.length
    ? `${g.students.length} alumnos · ${dates.length} día(s) · Por fecha: ${secs().filter((x) => x.type === 'asist' || x.type === 'sino').map((x) => x.name).join(' + ')}; las demás secciones aparte a la derecha`
    : 'Aún no hay registros en este grupo.';
  $('#preview-legend').innerHTML =
    STATUSES.map((k) => `<span class="summary-${k}">${escapeHtml(String(state.codes[k]))} = ${LABELS[k]}</span>`).join('') +
    (secs().some((x) => x.type === 'sino') ? READ.map((k) => `<span class="summary-${k === 'S' ? 'A' : 'F'}">${escapeHtml(String(state.codes[k]))} = ${k === 'S' ? 'Sí' : 'No'}</span>`).join('') : '');

  const d = $('#preview');
  d.showModal();
  // Columnas fijas (No. y Nombre) al desplazar horizontalmente.
  const firstRow = d.querySelectorAll('#preview-table tbody tr:first-child td');
  let left = 0;
  const lefts = [];
  firstRow.forEach((td, idx) => { if (idx <= g.nameCol) { lefts.push(left); left += td.offsetWidth; } });
  d.querySelectorAll('#preview-table tr').forEach((tr) =>
    [...tr.children].forEach((cell, idx) => { if (idx <= g.nameCol) cell.style.left = `${lefts[idx]}px`; })
  );
  // Desplaza hasta la fecha actual (o la última) para verla directo.
  const target = d.querySelector('#preview-table td.cur') || d.querySelector('#preview-table td.date-h:last-of-type');
  if (target) d.querySelector('.preview-scroll').scrollLeft = Math.max(0, target.offsetLeft - left - 40);
}

export async function doExport(btn, a = 'todo') {
  const txt = btn && btn.textContent;
  if (btn) { btn.disabled = true; btn.textContent = 'Generando…'; }
  try {
    const prep = preparar(state.groups, a);
    const out = await exportWorkbook(await loadExcel(), b64ToBuf(state.fileB64), prep.groups, state.codes, null, prep.periodos);
    await deliverFile(out, state.fileName.replace(/\.xlsx$/i, '') + (typeof a === 'number' ? ` - parcial ${a}.xlsx` : ' - asistencia.xlsx'));
    toast('Excel generado');
  } catch (err) {
    console.error(err);
    toast('Error al generar el Excel');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = txt; }
  }
}
