/*
 * Lógica de Excel: leer la lista del maestro y escribir las columnas de cada día.
 * Módulo ES: lo usan la app (Vite) y las pruebas de Node (npm test).
 *
 * Cada grupo tiene sus SECCIONES (g.sections), que el maestro puede agregar, quitar y renombrar.
 * Cada sección usa una PLANTILLA (type):
 *   asist  Asistencia A/F/R/J (siempre existe, id 'records')      g.records[fecha] = { fila: 'A' }
 *   sino   Sí / No, con mínimo de "Sí" para el 100 %               g[id][fecha] = { fila: 'S' | 'N' }
 *   trab   0 / 50 / 100, varios por día con título                 g[id][fecha] = [{ title, vals: { fila: 0|50|100 } }]
 *   disc   Disciplina: todos 10, solo se guarda el 0               g[id][fecha] = { fila: 0 }
 *   num    Calificación 0 a 100 (como examen)                      g[id][fecha] = { fila: 85 }
 *
 * Columnas del Excel, de izquierda a derecha (cada bloque separado por una columna vacía):
 *   por fecha: Asistencia + las secciones Sí/No · cada sección de trabajos · resumen de cada disciplina
 *   (Días con 0 y Fechas) · cada sección 0-100 (solo fechas con datos) · calificaciones.
 */

const STATUSES = ['A', 'F', 'R', 'J'];
const READ = ['S', 'N'];
const WORK = [0, 50, 100];
const DEFAULT_CODES = { A: 1, F: 0, R: 2, J: 3, S: 1, N: 0 };

// ---------- Secciones y plantillas ----------
const TYPES = {
  asist: { label: 'Asistencia (A / F / R / J)' },
  sino: { label: 'Sí / No (como lecturas o participaciones)' },
  trab: { label: '0 / 50 / 100 (como trabajos o tareas)' },
  disc: { label: 'Disciplina (todos 10, marcar 0)' },
  num: { label: 'Calificación de 0 a 100 (como examen)' },
};
const ASIST_ID = 'records';
const DEFAULT_SECTIONS = () => [
  { id: ASIST_ID, name: 'Asistencia', type: 'asist' },
  { id: 'lect', name: 'Participaciones', type: 'sino', min: 4 },
  { id: 'trab', name: 'Trabajos', type: 'trab' },
  { id: 'disc', name: 'Disciplina', type: 'disc' },
  { id: 'exam', name: 'Examen', type: 'num' },
];
// Alumnos activos: sin baja, o (si se da una fecha) dados de baja después de ese día.
const activos = (g, fecha = null) => g.students.filter((s) => !s.baja || (fecha != null && fecha < s.baja));
const sectionsOf = (g) => (g.sections && g.sections.length ? g.sections : DEFAULT_SECTIONS());
const newSectionId = () => `s_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
const secMin = (s) => Math.max(1, Math.round(Number(s && s.min) || 4));

function cellText(v) {
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    return '';
  }
  return String(v);
}

const clean = (s) => s.replace(/\s+/g, ' ').trim();
// Para comparar nombres: sin acentos, mayúsculas y espacios sencillos.
const norm = (s) => clean(String(s || '')).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const pad = (n) => String(n).padStart(2, '0');
const isoOf = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;

// Encabezado de una columna con fecha: celda fecha (versión anterior = asistencia)
// o texto "dd/mm/aa" + nombre de la sección (trabajos: "Nombre", "Nombre 2" o "Nombre: título").
function parseHeader(v, sections = DEFAULT_SECTIONS()) {
  const asist = sections.find((s) => s.type === 'asist');
  if (v instanceof Date && !isNaN(v)) return asist ? { date: isoOf(v), sec: asist.id } : null;
  if (v && typeof v === 'object' && v.result instanceof Date) return parseHeader(v.result, sections);
  const t = clean(cellText(v));
  const m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2}|\d{4})\b\s*(.*)$/);
  if (!m) return null;
  let y = Number(m[3]);
  if (y < 100) y += 2000;
  const date = `${y}-${pad(m[2])}-${pad(m[1])}`;
  const rest = clean(m[4] || '');
  if (!rest) return asist ? { date, sec: asist.id } : null;
  const R = norm(rest);
  const titleOf = () => (rest.includes(':') ? rest.slice(rest.indexOf(':') + 1).trim() : '');
  for (const s of sections) {
    const n = norm(s.name);
    if (s.type === 'trab') {
      if (new RegExp(`^${reEsc(n)}(\\s+\\d+)?(\\s*:.*)?$`).test(R)) return { date, sec: s.id, title: titleOf() };
    } else if (R === n) return { date, sec: s.id };
  }
  // Encabezados de versiones anteriores de la app.
  const first = (type) => sections.find((s) => s.type === type);
  if (/^ASIST/.test(R) && asist) return { date, sec: asist.id };
  if (/^(LECT|PARTICIP)/.test(R) && first('sino')) return { date, sec: first('sino').id };
  if (/^TRAB/.test(R) && first('trab')) return { date, sec: first('trab').id, title: titleOf() };
  if (/^EXAM/.test(R) && first('num')) return { date, sec: first('num').id };
  return null;
}
const parseDateHeader = (v) => { const h = parseHeader(v); return h ? h.date : null; };

// ---------- Trabajos: varios por día, cada uno con título ----------
const worksOf = (g, date, id = 'trab') => {
  const w = (g[id] || {})[date];
  if (Array.isArray(w)) return w;
  return w && typeof w === 'object' ? [{ title: '', vals: w }] : []; // formato anterior: un trabajo por día
};
const workUsed = (w) => !!w && (Object.keys(w.vals || {}).length > 0 || !!(w.title || '').trim());
// Hay datos de una sección en una fecha.
function hasData(g, sec, date) {
  if (typeof sec === 'string') sec = sectionsOf(g).find((s) => s.id === sec) || { id: sec, type: sec === 'trab' ? 'trab' : 'x' };
  if (sec.type === 'trab') return worksOf(g, date, sec.id).some((w) => Object.keys(w.vals || {}).length);
  return Object.keys(((g[sec.id] || {})[date]) || {}).length > 0;
}
// Pasa trabajos de la versión anterior ({fecha: {fila: v}}) al formato de varios trabajos.
function normalizeWorks(g, id = 'trab') {
  const t = g[id] || {};
  for (const d of Object.keys(t)) if (!Array.isArray(t[d])) t[d] = [{ title: '', vals: t[d] || {} }];
  g[id] = t;
  return g;
}
// Deja un grupo listo: secciones (las de siempre si no tiene) y un objeto de datos por sección.
function normalizeGroup(g, legacyGrading) {
  if (!g.sections || !g.sections.length) {
    g.sections = DEFAULT_SECTIONS();
    const lm = legacyGrading && legacyGrading.lectMin;
    if (lm) g.sections.find((s) => s.id === 'lect').min = lm;
  }
  if (!g.sections.some((s) => s.type === 'asist')) g.sections.unshift({ id: ASIST_ID, name: 'Asistencia', type: 'asist' });
  for (const s of g.sections) {
    g[s.id] = g[s.id] || {};
    if (s.type === 'trab') normalizeWorks(g, s.id);
    if (s.type === 'sino') s.min = secMin(s);
  }
  return g;
}

function valueFromCell(value, type, codes) {
  const t = clean(cellText(value)).toUpperCase();
  if (t === '') return null;
  if (type === 'trab') {
    const n = Number(t);
    return WORK.includes(n) ? n : null;
  }
  if (type === 'num') {
    const n = Number(t);
    return Number.isFinite(n) && n >= 0 && n <= 100 ? n : null;
  }
  const keys = type === 'sino' ? READ : STATUSES;
  return keys.find((k) => String(codes[k]).toUpperCase() === t) || null;
}

function findHeader(ws) {
  const maxRow = Math.min(ws.rowCount, 30);
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= Math.max(row.cellCount, 10); c++) {
      if (/^NOMBRE/i.test(clean(cellText(row.getCell(c).value)))) return { row: r, col: c };
    }
  }
  return null;
}

const LEGACY_GRADE = /^(ASISTENCIA|LECTURAS|TRABAJOS|DISCIPLINA|EXAMEN) \d+(\.\d+)?%$/;
// sectionsFor(nombreDeHoja) da las secciones de ese grupo (si ya existía), para reconocer sus columnas.
function parseWorkbook(wb, codes = DEFAULT_CODES, sectionsFor = null) {
  codes = Object.assign({}, DEFAULT_CODES, codes);
  const groups = [];
  wb.eachSheet((ws) => {
    const hdr = findHeader(ws);
    if (!hdr) return;
    const sections = JSON.parse(JSON.stringify((sectionsFor && sectionsFor(ws.name)) || DEFAULT_SECTIONS()));

    // Título: la celda que diga "GRUPO ..." arriba del encabezado, o el nombre de la hoja.
    let title = ws.name;
    for (let r = 1; r < hdr.row; r++) {
      const t = clean(cellText(ws.getCell(r, 1).value));
      if (/GRUPO/i.test(t)) title = t;
    }

    const students = [];
    for (let r = hdr.row + 1; r <= ws.rowCount; r++) {
      const name = clean(cellText(ws.getCell(r, hdr.col).value));
      if (!name) continue;
      const no = hdr.col > 1 ? clean(cellText(ws.getCell(r, hdr.col - 1).value)) : '';
      students.push({ row: r, no, name });
    }
    if (!students.length) return;

    // Columnas escritas por la app (por si suben un Excel exportado): se recuperan los datos.
    const data = {};
    sections.forEach((s) => (data[s.id] = {}));
    const dateCols = [];
    const lastCol = Math.max(ws.columnCount, hdr.col);
    const discs = sections.filter((s) => s.type === 'disc');
    for (let c = hdr.col + 1; c <= lastCol; c++) {
      const ht = norm(cellText(ws.getCell(hdr.row, c).value));
      // Calificaciones: se reconocen para poder reescribirlas.
      if (/^CALIFICACION SOBRE \d+$/.test(ht) || /^PARCIAL \d+$/.test(ht) || /^SEMESTRAL SOBRE \d+$/.test(ht) || LEGACY_GRADE.test(ht) ||
        sections.some((s) => new RegExp(`^${reEsc(norm(s.name))} \\d+(\\.\\d+)?%$`).test(ht))) {
        dateCols.push(c);
        continue;
      }
      // Resumen de disciplina: de la columna "Fechas" se recuperan los días con 0.
      const ds = discs.find((s) => ht === `${norm(s.name)} DIAS CON 0` || ht === `${norm(s.name)} FECHAS`) ||
        (/^DISCIPLINA (DIAS CON 0|FECHAS)$/.test(ht) ? discs[0] || { id: null } : null);
      if (ds) {
        dateCols.push(c);
        if (ds.id && /FECHAS$/.test(ht)) {
          for (const s of students) {
            const t = cellText(ws.getCell(s.row, c).value);
            for (const m of t.matchAll(/(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})/g)) {
              const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
              const d = `${y}-${pad(m[2])}-${pad(m[1])}`;
              (data[ds.id][d] = data[ds.id][d] || {})[s.row] = 0;
            }
          }
        }
        continue;
      }
      const h = parseHeader(ws.getCell(hdr.row, c).value, sections);
      if (!h) continue;
      const sec = sections.find((s) => s.id === h.sec);
      dateCols.push(c);
      let day;
      if (sec.type === 'trab') {
        // Cada columna de trabajo es un trabajo distinto de esa fecha.
        const w = { title: h.title || '', vals: {} };
        (data[sec.id][h.date] = data[sec.id][h.date] || []).push(w);
        day = w.vals;
      } else day = (data[sec.id][h.date] = data[sec.id][h.date] || {});
      for (const s of students) {
        const v = valueFromCell(ws.getCell(s.row, c).value, sec.type, codes);
        if (v != null) day[s.row] = v;
      }
    }

    // Quita días sin ningún dato.
    for (const sec of sections) {
      const k = sec.id;
      for (const d of Object.keys(data[k])) {
        if (sec.type === 'trab') {
          data[k][d] = data[k][d].filter(workUsed);
          if (!data[k][d].length) delete data[k][d];
        } else if (!Object.keys(data[k][d]).length) delete data[k][d];
      }
    }

    // Primera columna libre después del nombre: encabezado vacío y sin datos de alumnos.
    let startCol;
    if (dateCols.length) {
      startCol = Math.min(...dateCols);
    } else {
      startCol = hdr.col + 1;
      const used = (c) =>
        clean(cellText(ws.getCell(hdr.row, c).value)) !== '' ||
        students.some((s) => clean(cellText(ws.getCell(s.row, c).value)) !== '');
      while (used(startCol)) startCol++;
    }

    groups.push(Object.assign({ sheet: ws.name, title, headerRow: hdr.row, nameCol: hdr.col, startCol, dateCols, students, sections }, data));
  });
  return groups;
}

function codeValue(key, codes) {
  const v = codes[key];
  if (v === '' || v == null) return null;
  const n = Number(v);
  return Number.isFinite(n) && String(v).trim() !== '' ? n : String(v);
}

const shortDate = (iso) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y.slice(-2)}`; };
// Fechas con datos de una sección, ordenadas.
const datesOf = (g, sec) => Object.keys(g[sec.id] || {}).filter((d) => hasData(g, sec, d)).sort();
// Fechas con algún dato de asistencia, Sí/No o trabajos (las que llevan columna por día).
function groupDates(g) {
  const set = new Set();
  sectionsOf(g).filter((s) => ['asist', 'sino', 'trab'].includes(s.type)).forEach((s) => datesOf(g, s).forEach((d) => set.add(d)));
  return [...set].sort();
}
// Fechas (ordenadas) en que el alumno sacó 0 en una sección de disciplina.
function discDates(g, row, id = 'disc') {
  return Object.keys(g[id] || {}).filter((d) => g[id][d][row] === 0).sort();
}

// ---------- Calificaciones ----------
const DEFAULT_GRADING = { w: {}, retardo: 100, escala: 10 };
const round1 = (n) => Math.round(n * 10) / 10;
const avg = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : null);
const weightOf = (grading, s) => Number(((grading || {}).w || {})[s.id]) || 0;
const gradingOn = (grading, g) => (g ? sectionsOf(g) : Object.keys((grading || {}).w || {}).map((id) => ({ id })))
  .some((s) => weightOf(grading, s) > 0);
// Días en que el alumno estuvo en clase: A o R en asistencia (o, sin asistencia, días con algún dato suyo).
function presentDays(g, row) {
  const secs = sectionsOf(g);
  const asist = secs.find((s) => s.type === 'asist');
  const set = new Set();
  if (asist) Object.keys(g[asist.id] || {}).forEach((d) => { const v = g[asist.id][d][row]; if (v === 'A' || v === 'R') set.add(d); });
  return set;
}

// Porcentaje (0-100) de cada sección para un alumno; null si aún no hay datos.
// Solo cuenta lo capturado: un día sin captura para ese alumno no suma ni resta.
function gradeParts(g, row, grading) {
  grading = Object.assign({}, DEFAULT_GRADING, grading);
  const out = {};
  const vals = (id) => Object.keys(g[id] || {}).map((d) => [d, g[id][d][row]]).filter(([, v]) => v != null);
  for (const s of sectionsOf(g)) {
    if (s.type === 'asist') {
      // Asistencia y Justificado = 100 %, Retardo = lo que diga el maestro, Falta = 0 %.
      const sc = { A: 100, J: 100, R: Number(grading.retardo) || 0, F: 0 };
      out[s.id] = avg(vals(s.id).map(([, v]) => sc[v]).filter((v) => v != null));
    } else if (s.type === 'sino') {
      // Con el mínimo de "Sí" (o más) tiene el 100 %; debajo, proporcional.
      const v = vals(s.id);
      const si = v.filter(([, x]) => x === 'S').length;
      out[s.id] = v.length ? (100 * Math.min(si, secMin(s))) / secMin(s) : null;
    } else if (s.type === 'trab') {
      out[s.id] = avg(Object.keys(g[s.id] || {}).flatMap((d) => worksOf(g, d, s.id).map((w) => (w.vals || {})[row]))
        .filter((v) => v != null).map(Number));
    } else if (s.type === 'disc') {
      // Días en clase (A o R) o con reporte; 10 salvo los días con 0.
      const present = presentDays(g, row);
      const zeros = discDates(g, row, s.id);
      zeros.forEach((d) => present.add(d));
      out[s.id] = present.size ? (100 * (present.size - zeros.length)) / present.size : null;
    } else if (s.type === 'num') {
      out[s.id] = avg(vals(s.id).map(([, v]) => Number(v)));
    }
  }
  return out;
}

// Calificación final en la escala elegida (10 o 100): el alumno empieza con todo y solo se le descuenta
// lo que ya perdió. Una sección sin datos todavía no descuenta (cuenta completa); p. ej. participaciones
// al 50 % con valor de 10 % = −5 → 95. Con todas las secciones capturadas es el promedio ponderado normal.
// Sin ningún dato en las secciones con porcentaje no hay calificación (null).
function finalGrade(g, row, grading) {
  grading = Object.assign({}, DEFAULT_GRADING, grading);
  const parts = gradeParts(g, row, grading);
  let sum = 0;
  let wsum = 0;
  let any = false;
  for (const s of sectionsOf(g)) {
    const w = weightOf(grading, s);
    if (w <= 0) continue;
    wsum += w;
    if (parts[s.id] != null) any = true;
    sum += w * (parts[s.id] != null ? parts[s.id] : 100);
  }
  const final = wsum && any ? round1(((sum / wsum) * (Number(grading.escala) || 10)) / 100) : null;
  return { parts, final };
}

// ---------- Acomodo de columnas ----------
// Lista de columnas que escribe la app: { col, kind, sec, date, work, block, first, last, label }.
// kind: 'day' (asistencia y Sí/No por fecha) · 'trab' · 'disc-count' · 'disc-dates' · 'num' · 'grade' (key: id o 'final').
function sheetLayout(g, grading, periodos = null) {
  grading = grading || g.grading;
  const secs = sectionsOf(g);
  const out = [];
  let col = g.startCol;
  const gap = () => { if (out.length) col++; };
  const close = () => { if (out.length) out[out.length - 1].last = true; };
  // 1. Por fecha: asistencia + secciones Sí/No (gris).
  const dayS = secs.filter((s) => s.type === 'asist' || s.type === 'sino');
  const dayDates = [...new Set(dayS.flatMap((s) => datesOf(g, s)))].sort();
  const asist = dayS.find((s) => s.type === 'asist');
  dayDates.forEach((date) => {
    let first = true;
    dayS.forEach((s) => {
      if (s.type === 'sino' && !hasData(g, s, date)) return;
      out.push({ col: col++, kind: 'day', sec: s, date, block: 'day', first, label: `${shortDate(date)}\n${s.name}` });
      first = false;
    });
  });
  close();
  // 2. Cada sección de trabajos: una columna por trabajo con calificaciones.
  secs.filter((s) => s.type === 'trab').forEach((s) => {
    const items = [];
    datesOf(g, s).forEach((date) => {
      const ws = worksOf(g, date, s.id);
      const used = ws.filter((w) => Object.keys(w.vals || {}).length);
      ws.forEach((w, work) => {
        if (!Object.keys(w.vals || {}).length) return;
        const n = used.indexOf(w) + 1;
        const title = (w.title || '').trim();
        items.push({ kind: 'trab', sec: s, date, work, label: `${shortDate(date)}\n${title ? `${s.name}: ${title}` : n > 1 ? `${s.name} ${n}` : s.name}`, titled: !!title });
      });
    });
    if (!items.length) return;
    gap();
    items.forEach((x, i) => out.push(Object.assign(x, { col: col++, block: `trab:${s.id}`, first: i === 0 })));
    close();
  });
  // 3. Disciplina: 2 columnas de resumen por sección (si hay cualquier dato en el grupo).
  const anyData = out.length || secs.some((s) => (s.type === 'disc' || s.type === 'num') && datesOf(g, s).length);
  if (anyData) {
    secs.filter((s) => s.type === 'disc').forEach((s) => {
      gap();
      out.push({ col: col++, kind: 'disc-count', sec: s, block: `disc:${s.id}`, first: true, label: `${s.name}\nDías con 0` });
      out.push({ col: col++, kind: 'disc-dates', sec: s, block: `disc:${s.id}`, last: true, label: `${s.name}\nFechas` });
    });
  }
  // 4. Secciones de 0 a 100: una columna por fecha con datos.
  secs.filter((s) => s.type === 'num').forEach((s) => {
    const ds = datesOf(g, s);
    if (!ds.length) return;
    gap();
    ds.forEach((date, i) => out.push({ col: col++, kind: 'num', sec: s, date, block: `num:${s.id}`, first: i === 0, label: `${shortDate(date)}\n${s.name}` }));
    close();
  });
  // 5. Calificaciones: una por sección con porcentaje y la final.
  if (gradingOn(grading, g)) {
    gap();
    const ps = secs.filter((s) => weightOf(grading, s) > 0);
    ps.forEach((s, i) => out.push({ col: col++, kind: 'grade', key: s.id, sec: s, block: 'grade', first: i === 0, label: `${s.name}\n${weightOf(grading, s)}%` }));
    out.push({ col: col++, kind: 'grade', key: 'final', block: 'grade', first: true, last: true, label: `Calificación\nsobre ${Number(grading.escala) || 10}` });
    // Con parciales definidos: la calificación de cada parcial y la semestral (cerradas si ya se cerraron).
    if (periodosListos(periodos)) {
      out[out.length - 1].last = false;
      periodos.parciales.forEach((p, i) => out.push({ col: col++, kind: 'grade', key: `p${p.numero}`, parcial: p, block: 'grade', first: i === 0, label: `Parcial ${p.numero}` }));
      out.push({ col: col++, kind: 'grade', key: 'semestral', block: 'grade', first: true, last: true, label: `Semestral\nsobre ${Number(grading.escala) || 10}` });
    }
  }
  return out;
}

// Escribe los datos en una copia del Excel original y devuelve el buffer .xlsx
async function exportWorkbook(ExcelJS, originalBuffer, groups, codes = DEFAULT_CODES, grading = null, periodos = null) {
  codes = Object.assign({}, DEFAULT_CODES, codes);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(originalBuffer);
  const thin = { style: 'thin' };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const thick = { style: 'medium', color: { argb: 'FF555555' } };
  const grayFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
  const finalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

  for (const g of groups) {
    const ws = wb.getWorksheet(g.sheet);
    if (!ws) continue;
    sectionsOf(g).filter((s) => s.type === 'trab').forEach((s) => normalizeWorks(g, s.id));

    // Limpia las columnas que la app había escrito (se reescriben ordenadas).
    for (const c of g.dateCols || []) {
      for (const r of [g.headerRow, ...g.students.map((s) => s.row)]) {
        const cell = ws.getCell(r, c);
        cell.value = null;
        cell.style = {};
      }
    }

    // Alumnos agregados en la app: al final de la lista con el formato del último alumno original.
    const added = g.students.filter((s) => s.added);
    if (added.length) {
      const original = g.students.filter((s) => !s.added);
      const tmplRow = original.length ? Math.max(...original.map((s) => s.row)) : g.headerRow;
      const lastCol = ws.columnCount;
      for (const s of added) {
        for (let c = 1; c <= lastCol; c++) {
          ws.getCell(s.row, c).style = JSON.parse(JSON.stringify(ws.getCell(tmplRow, c).style || {}));
        }
        if (ws.getRow(tmplRow).height) ws.getRow(s.row).height = ws.getRow(tmplRow).height;
        if (g.nameCol > 1 && s.no !== '') {
          const n = Number(s.no);
          ws.getCell(s.row, g.nameCol - 1).value = Number.isFinite(n) ? n : s.no;
        }
        ws.getCell(s.row, g.nameCol).value = s.name;
      }
    }

    const gg = g.grading || grading;
    const layout = sheetLayout(g, gg, periodos);
    const headerSrc = ws.getCell(g.headerRow, g.nameCol);
    if (layout.length) {
      const hr = ws.getRow(g.headerRow);
      hr.height = Math.max(hr.height || 15, layout.some((x) => x.titled) ? 45 : 30);
    }
    const finals = {};
    const gradeOf = (row) => (finals[row] = finals[row] || finalGrade(g, row, gg));
    layout.forEach((x) => {
      const { col, kind, sec, date } = x;
      const isFinal = kind === 'grade' && x.key === 'final';
      const cellBorder = Object.assign({}, border, x.first ? { left: thick } : {}, x.last ? { right: thick } : {});
      const fill = kind === 'day' && sec.type === 'sino' ? grayFill : isFinal ? finalFill : null;
      // Estilo completo y propio de cada celda (ExcelJS comparte objetos de estilo entre celdas).
      const h = ws.getCell(g.headerRow, col);
      h.value = x.label;
      h.style = {
        font: Object.assign({}, headerSrc.font, { bold: true }),
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: cellBorder,
        ...(fill ? { fill } : {}),
      };
      ws.getColumn(col).width = kind === 'disc-dates' ? 24 : x.titled ? 16 : isFinal ? 13 : 11;
      for (const s of g.students) {
        let v = null;
        if (kind === 'day') {
          const raw = ((g[sec.id] || {})[date] || {})[s.row];
          v = raw == null ? null : codeValue(raw, codes);
        } else if (kind === 'trab') {
          const raw = (worksOf(g, date, sec.id)[x.work].vals || {})[s.row];
          v = raw == null ? null : Number(raw);
        } else if (kind === 'num') {
          const raw = ((g[sec.id] || {})[date] || {})[s.row];
          v = raw == null ? null : Number(raw);
        } else if (kind.startsWith('disc')) {
          const z = discDates(g, s.row, sec.id);
          v = kind === 'disc-count' ? z.length : (z.map(shortDate).join(', ') || null);
        } else if (kind === 'grade' && x.parcial) {
          const c = cierreVigente(g, 'parcial', x.parcial.numero);
          const raw = c && c.calificaciones[s.row] ? c.calificaciones[s.row].final : finalGrade(enPeriodo(g, x.parcial.inicio, x.parcial.fin), s.row, gg).final;
          v = raw == null ? null : round1(raw);
        } else if (kind === 'grade' && x.key === 'semestral') {
          const c = cierreVigente(g, 'semestre', null);
          v = c && c.calificaciones[s.row] ? c.calificaciones[s.row].reportada : semestral(g, s.row, gg, periodos, g.semestral).reportada;
          if (v == null) v = null;
        } else if (kind === 'grade') {
          const r = gradeOf(s.row);
          const raw = isFinal ? r.final : r.parts[x.key];
          v = raw == null ? null : round1(raw);
        }
        const cell = ws.getCell(s.row, col);
        cell.value = v;
        cell.style = {
          font: Object.assign({}, ws.getCell(s.row, g.nameCol).font || {}, isFinal ? { bold: true } : {}),
          alignment: { horizontal: kind === 'disc-dates' ? 'left' : 'center', vertical: 'middle', wrapText: kind === 'disc-dates' },
          border: cellBorder,
          ...(fill ? { fill } : {}),
        };
      }
    });
  }
  return wb.xlsx.writeBuffer();
}

// ---------- Resumen de un alumno ----------
// Todo lo capturado de un alumno, día por día y por sección, más su calificación. Lo usan la pantalla
// "Resumen por alumno" y el Excel individual (para quejas del alumno o de los padres).
const STATUS_LABEL = { A: 'Asistencia', F: 'Falta', R: 'Retardo', J: 'Justificado' };
const WEEKDAY = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'];
const longDate = (iso) => {
  const [y, m, d] = iso.split('-').map(Number);
  return `${WEEKDAY[new Date(y, m - 1, d).getDay()]} ${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
};
const workLabel = (w, i, n) => (w.title ? w.title : n > 1 ? `Trabajo ${i + 1}` : 'Trabajo');

function studentReport(g, row, grading) {
  grading = Object.assign({}, DEFAULT_GRADING, grading);
  const secs = sectionsOf(g);
  const student = g.students.find((s) => s.row === row) || null;
  const at = (id, d) => ((g[id] || {})[d] || {})[row];
  const present = presentDays(g, row);
  const dates = new Set();
  for (const s of secs) {
    if (s.type === 'trab') Object.keys(g[s.id] || {}).forEach((d) => { if (worksOf(g, d, s.id).some((w) => (w.vals || {})[row] != null)) dates.add(d); });
    else Object.keys(g[s.id] || {}).forEach((d) => { if (at(s.id, d) != null) dates.add(d); });
  }
  const counts = { A: 0, F: 0, R: 0, J: 0 };
  const stats = {};
  secs.forEach((s) => (stats[s.id] = { S: 0, N: 0, works: 0, list: [] }));
  // Celda de cada sección en un día: texto para el Excel, texto corto para la pantalla y si es "mala" (rojo).
  const days = [...dates].sort().map((date) => {
    const cells = {};
    for (const s of secs) {
      let c = null;
      if (s.type === 'asist') {
        const v = at(s.id, date);
        if (v) { counts[v]++; c = { text: STATUS_LABEL[v], short: v, tone: v, bad: v === 'F' }; }
      } else if (s.type === 'sino') {
        const v = at(s.id, date);
        if (v) { stats[s.id][v]++; c = { text: v === 'S' ? 'Sí' : 'No', short: v === 'S' ? 'Sí' : 'No', tone: v === 'S' ? 'A' : 'F', bad: v === 'N' }; }
      } else if (s.type === 'trab') {
        const ws = worksOf(g, date, s.id);
        const list = ws.map((w, i) => ({ title: workLabel(w, i, ws.length), v: (w.vals || {})[row] })).filter((w) => w.v != null);
        if (list.length) {
          stats[s.id].works += list.length;
          const t = list.map((w) => `${w.title}: ${Number(w.v)}`).join('\n');
          c = { text: t, short: t, bad: list.some((w) => Number(w.v) === 0), list };
        }
      } else if (s.type === 'disc') {
        // 0 si tuvo reporte; 10 si estuvo en clase sin reporte; vacío si no vino.
        const v = at(s.id, date) === 0 ? 0 : present.has(date) ? 10 : null;
        if (v != null) c = { text: v, short: v, bad: v === 0, tone: v === 0 ? 'F' : '' };
      } else if (s.type === 'num') {
        const v = at(s.id, date);
        if (v != null) { stats[s.id].list.push({ date, v: Number(v) }); c = { text: Number(v), short: Number(v) }; }
      }
      if (c) cells[s.id] = c;
    }
    return { date, label: longDate(date), cells };
  });
  const { parts, final } = finalGrade(g, row, grading);
  const detail = {};
  for (const s of secs) {
    const st = stats[s.id];
    if (s.type === 'asist') detail[s.id] = `Asistencias ${counts.A} · Faltas ${counts.F} · Retardos ${counts.R} · Justificados ${counts.J}`;
    else if (s.type === 'sino') detail[s.id] = st.S + st.N ? `Sí ${st.S} de ${st.S + st.N} · mínimo para el 100 %: ${secMin(s)}` : 'Sin datos';
    else if (s.type === 'trab') detail[s.id] = st.works ? `${st.works} calificación(es)` : 'Sin datos';
    else if (s.type === 'disc') { const z = discDates(g, row, s.id); detail[s.id] = z.length ? `${z.length} día(s) con 0: ${z.map(shortDate).join(', ')}` : 'Sin días con 0'; }
    else if (s.type === 'num') detail[s.id] = st.list.length ? st.list.map((e) => `${shortDate(e.date)}: ${e.v}`).join(' · ') : 'Sin datos';
  }
  return { student, sections: secs, days, counts, parts, final, detail, grading, graded: gradingOn(grading, g) };
}

// Excel nuevo (no es copia del original) con el concentrado de un solo alumno.
function exportStudent(ExcelJS, g, row, grading) {
  const r = studentReport(g, row, grading);
  const s = r.student;
  const secs = r.sections;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Resumen', { views: [{ showGridLines: false }] });
  const nCols = Math.max(4, secs.length + 1);
  ws.columns = Array.from({ length: nCols }, (_, i) => ({ width: i === 0 ? 18 : i === 3 && nCols === 4 ? 34 : 16 }));
  const thin = { style: 'thin', color: { argb: 'FF999999' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const HEAD = fill('FFD9E2F3');
  const RED = fill('FFF8D7D7');
  const YEL = fill('FFFFF2CC');
  const style = (cell, o = {}) => {
    cell.style = {
      font: { name: 'Calibri', size: o.size || 11, bold: !!o.bold, color: o.color ? { argb: o.color } : undefined },
      alignment: { vertical: 'middle', horizontal: o.align || 'left', wrapText: true },
      ...(o.box ? { border } : {}),
      ...(o.fill ? { fill: o.fill } : {}),
    };
  };
  const put = (r0, c, v, o) => { const cell = ws.getCell(r0, c); cell.value = v; style(cell, o); return cell; };

  let y = 1;
  put(y, 1, 'Resumen del alumno', { bold: true, size: 16 });
  ws.mergeCells(y, 1, y, 4);
  y += 2;
  const today = new Date();
  const info = [['Alumno', s ? s.name : ''], ['No. de lista', s ? s.no : ''], ['Grupo', g.title ? `${g.sheet} — ${g.title}` : g.sheet], ...(g.subject ? [['Materia', g.subject]] : []),
    ['Generado', longDate(`${today.getFullYear()}-${today.getMonth() + 1}-${today.getDate()}`)]];
  info.forEach(([k, v]) => { put(y, 1, k, { bold: true }); ws.mergeCells(y, 2, y, 4); put(y, 2, v == null ? '' : String(v)); y++; });
  y++;

  // Calificación por sección
  put(y, 1, 'Calificación', { bold: true, size: 13 });
  y++;
  ['Sección', 'Vale', 'Resultado', 'Detalle'].forEach((t, i) => put(y, i + 1, t, { bold: true, box: true, fill: HEAD, align: i === 3 ? 'left' : 'center' }));
  if (nCols > 4) ws.mergeCells(y, 4, y, nCols);
  y++;
  for (const p of secs) {
    const w = weightOf(r.grading, p);
    const v = r.parts[p.id];
    put(y, 1, p.name, { box: true, bold: true });
    put(y, 2, w ? `${w} %` : '—', { box: true, align: 'center' });
    put(y, 3, v == null ? 'Sin datos' : `${round1(v)} %`, { box: true, align: 'center' });
    put(y, 4, r.detail[p.id], { box: true });
    if (nCols > 4) ws.mergeCells(y, 4, y, nCols);
    y++;
  }
  put(y, 1, `Calificación final (sobre ${r.grading.escala})`, { bold: true, box: true, fill: YEL });
  ws.mergeCells(y, 1, y, 2);
  put(y, 3, r.final == null ? (r.graded ? 'Sin datos' : 'Sin porcentajes') : r.final, { bold: true, box: true, fill: YEL, align: 'center' });
  put(y, 4, r.graded ? 'Empieza con la calificación completa; las secciones sin datos todavía no descuentan.' : 'El maestro aún no define cuánto vale cada sección.', { box: true, fill: YEL });
  if (nCols > 4) ws.mergeCells(y, 4, y, nCols);
  y += 2;

  // Detalle por día: una columna por sección.
  put(y, 1, 'Detalle por día', { bold: true, size: 13 });
  y++;
  ['Fecha', ...secs.map((p) => p.name)].forEach((t, i) => put(y, i + 1, t, { bold: true, box: true, fill: HEAD, align: 'center' }));
  y++;
  if (!r.days.length) {
    ws.mergeCells(y, 1, y, secs.length + 1);
    put(y, 1, 'Aún no hay datos capturados de este alumno.', { box: true });
  }
  for (const d of r.days) {
    put(y, 1, d.label, { box: true });
    secs.forEach((p, i) => {
      const c = d.cells[p.id];
      put(y, i + 2, c ? c.text : '', { box: true, align: p.type === 'trab' ? 'left' : 'center', fill: c && c.bad ? RED : undefined });
    });
    y++;
  }
  return wb.xlsx.writeBuffer();
}

// ---------- Resumen del grupo (semáforo) ----------
// Para el director: promedios, faltas y alumnos en riesgo. El estado de cada alumno sale de lo peor entre
// su calificación y su asistencia: rojo si está reprobando o falta mucho, amarillo si está cerca, verde si va bien.
const RISK = { failPct: 60, warnPct: 80, attRed: 80, attYellow: 90 };
const STATUS_TXT = { rojo: 'Rojo', amarillo: 'Amarillo', verde: 'Verde', none: 'Sin datos' };

function groupReport(g, grading) {
  grading = Object.assign({}, DEFAULT_GRADING, grading);
  const escala = Number(grading.escala) || 10;
  const graded = gradingOn(grading, g);
  let present = 0;
  let taken = 0;
  const rows = activos(g).map((s) => {
    const c = { A: 0, F: 0, R: 0, J: 0 };
    Object.values(g[ASIST_ID] || {}).forEach((day) => { const v = day[s.row]; if (v in c) c[v]++; });
    const total = c.A + c.F + c.R + c.J;
    present += total - c.F;
    taken += total;
    const att = total ? (100 * (total - c.F)) / total : null;
    const { parts, final } = finalGrade(g, s.row, grading);
    const fp = graded && final != null ? (100 * final) / escala : null;
    const reasons = [];
    let status = 'none';
    const worse = (st) => { const o = ['none', 'verde', 'amarillo', 'rojo']; if (o.indexOf(st) > o.indexOf(status)) status = st; };
    if (fp != null) {
      if (fp < RISK.failPct) { worse('rojo'); reasons.push(`Calificación ${final} (reprobando)`); }
      else if (fp < RISK.warnPct) { worse('amarillo'); reasons.push(`Calificación ${final}`); }
      else worse('verde');
    }
    if (att != null) {
      if (att < RISK.attRed) { worse('rojo'); reasons.push(`Asistencia ${round1(att)} %`); }
      else if (att < RISK.attYellow) { worse('amarillo'); reasons.push(`Asistencia ${round1(att)} %`); }
      else worse('verde');
    }
    return { student: s, counts: c, days: total, att, parts, final, status, reasons };
  });
  const finals = rows.map((r) => r.final).filter((v) => v != null);
  const count = (st) => rows.filter((r) => r.status === st).length;
  return {
    rows, graded, escala, grading,
    stats: {
      alumnos: activos(g).length,
      asistencia: taken ? (100 * present) / taken : null,
      promedio: graded && finals.length ? round1(finals.reduce((a, b) => a + b, 0) / finals.length) : null,
      faltas: rows.reduce((n, r) => n + r.counts.F, 0),
      rojo: count('rojo'), amarillo: count('amarillo'), verde: count('verde'), none: count('none'),
    },
    criteria: [
      `Verde: calificación de ${round1((escala * RISK.warnPct) / 100)} o más y asistencia de ${RISK.attYellow} % o más.`,
      `Amarillo: calificación de ${round1((escala * RISK.failPct) / 100)} a menos de ${round1((escala * RISK.warnPct) / 100)}, o asistencia de ${RISK.attRed} % a menos de ${RISK.attYellow} %.`,
      `Rojo (en riesgo): calificación menor a ${round1((escala * RISK.failPct) / 100)} o asistencia menor a ${RISK.attRed} %.`,
    ],
  };
}

function exportGroupReport(ExcelJS, g, grading) {
  const r = groupReport(g, grading);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Resumen del grupo', { views: [{ showGridLines: false }] });
  const parts = sectionsOf(g).filter((p) => weightOf(r.grading, p) > 0);
  const head = ['No.', 'Nombre', 'Asistencias', 'Faltas', 'Retardos', 'Justificados', '% Asistencia',
    ...parts.map((p) => `${p.name}\n${weightOf(r.grading, p)}%`), `Calificación\nsobre ${r.escala}`, 'Estado', 'Motivo'];
  ws.columns = head.map((h, i) => ({ width: i === 0 ? 6 : i === 1 ? 36 : h === 'Motivo' ? 34 : 12 }));
  const thin = { style: 'thin', color: { argb: 'FF999999' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const SOFT = { verde: 'FFD9EAD3', amarillo: 'FFFFF2CC', rojo: 'FFF4CCCC', none: 'FFF2F2F2' };
  const STRONG = { verde: 'FF38761D', amarillo: 'FFBF9000', rojo: 'FFCC0000', none: 'FF999999' };
  const put = (row, col, v, o = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = v;
    cell.style = {
      font: { name: 'Calibri', size: o.size || 11, bold: !!o.bold, color: o.color ? { argb: o.color } : undefined },
      alignment: { vertical: 'middle', horizontal: o.align || 'left', wrapText: true },
      ...(o.box ? { border } : {}),
      ...(o.fill ? { fill: fill(o.fill) } : {}),
    };
    return cell;
  };
  let y = 1;
  put(y, 1, 'Resumen del grupo', { bold: true, size: 16 });
  ws.mergeCells(y, 1, y, 6);
  ws.getRow(y).height = 24;
  y += 2;
  const today = new Date();
  const iso = `${today.getFullYear()}-${today.getMonth() + 1}-${today.getDate()}`;
  [['Grupo', g.title ? `${g.sheet} — ${g.title}` : g.sheet], ...(g.subject ? [['Materia', g.subject]] : []), ['Generado', longDate(iso)]]
    .forEach(([k, v]) => { put(y, 1, k, { bold: true }); ws.mergeCells(y, 2, y, 4); put(y, 2, v); y++; });
  y++;
  // Indicadores
  const st = r.stats;
  const tiles = [
    ['Alumnos', st.alumnos, 'FFD9E2F3'],
    ['Asistencia promedio', st.asistencia == null ? 'Sin datos' : `${round1(st.asistencia)} %`, 'FFD9E2F3'],
    ['Promedio general', st.promedio == null ? (r.graded ? 'Sin datos' : 'Sin porcentajes') : st.promedio, 'FFD9E2F3'],
    ['Faltas en total', st.faltas, 'FFD9E2F3'],
    ['En riesgo (rojo)', st.rojo, SOFT.rojo],
    ['En observación (amarillo)', st.amarillo, SOFT.amarillo],
    ['Bien (verde)', st.verde, SOFT.verde],
  ];
  tiles.forEach(([k, v, f]) => {
    put(y, 1, k, { bold: true, box: true, fill: f });
    ws.mergeCells(y, 1, y, 2);
    put(y, 3, v, { bold: true, box: true, fill: f, align: 'center', size: 12 });
    y++;
  });
  y++;
  put(y, 1, 'Criterios del semáforo', { bold: true });
  ws.mergeCells(y, 1, y, 4);
  y++;
  r.criteria.forEach((t, i) => {
    ws.mergeCells(y, 1, y, Math.min(head.length, 8));
    put(y, 1, t, { color: STRONG[['verde', 'amarillo', 'rojo'][i]] });
    y++;
  });
  y++;
  // Tabla por alumno
  const hy = y;
  head.forEach((h, i) => put(y, i + 1, h, { bold: true, box: true, fill: 'FFD9E2F3', align: 'center' }));
  ws.getRow(y).height = 32;
  y++;
  for (const row of r.rows) {
    const f = SOFT[row.status];
    const vals = [row.student.no, row.student.name, row.counts.A, row.counts.F, row.counts.R, row.counts.J,
      row.att == null ? '' : round1(row.att),
      ...parts.map((p) => (row.parts[p.id] == null ? '' : round1(row.parts[p.id]))),
      row.final == null ? '' : row.final, STATUS_TXT[row.status], row.reasons.join(' · ')];
    vals.forEach((v, i) => {
      const isState = i === vals.length - 2;
      put(y, i + 1, v, {
        box: true,
        fill: isState ? STRONG[row.status] : f,
        color: isState ? 'FFFFFFFF' : undefined,
        bold: isState || i === vals.length - 3,
        align: i === 1 || i === vals.length - 1 ? 'left' : 'center',
      });
    });
    y++;
  }
  ws.autoFilter = { from: { row: hy, column: 1 }, to: { row: y - 1, column: head.length } };
  return wb.xlsx.writeBuffer();
}


// ---------- Parciales y semestre ----------
// periodos = { semestre, inicio, fin, parciales: [{ numero, inicio, fin }] } (fechas ISO, inclusivas).
// Vista del grupo solo con las fechas de un periodo (para calcular la calificación de un parcial).
function enPeriodo(g, desde, hasta) {
  const v = Object.assign({}, g);
  for (const s of sectionsOf(g)) {
    const datos = g[s.id] || {};
    v[s.id] = {};
    for (const d of Object.keys(datos)) if ((!desde || d >= desde) && (!hasta || d <= hasta)) v[s.id][d] = datos[d];
  }
  return v;
}
const periodosListos = (p) => !!(p && Array.isArray(p.parciales) && p.parciales.length && p.parciales.every((x) => x.inicio && x.fin));
// Parcial al que pertenece una fecha (número) o null.
const parcialDe = (p, fecha) => (periodosListos(p) ? (p.parciales.find((x) => fecha >= x.inicio && fecha <= x.fin) || {}).numero || null : null);

// Cierre vigente (no reabierto) de un parcial o del semestre.
const cierreVigente = (g, tipo, n) => (g.cierres || []).filter((c) => c.tipo === tipo && (c.parcial ?? null) === n && !c.reabierto_en).at(-1) || null;
const SEMESTRAL = { modo: 'promedio', pesoParciales: 60, pesoFinal: 40, exentoDesde: null, minimo: null, redondeo: 'decimal', reprobado5: false };
// Calificación semestral de un alumno. config: SEMESTRAL + final: { fila: 0-100 }.
function semestral(g, row, grading, periodos, config) {
  grading = Object.assign({}, DEFAULT_GRADING, grading);
  const c = Object.assign({}, SEMESTRAL, config);
  const escala = Number(grading.escala) || 10;
  const parciales = periodosListos(periodos)
    ? periodos.parciales.map((x) => {
      // Un parcial cerrado vale lo que quedó al cerrarlo.
      const c = cierreVigente(g, 'parcial', x.numero);
      return c && c.calificaciones[row] ? c.calificaciones[row].final : finalGrade(enPeriodo(g, x.inicio, x.fin), row, grading).final;
    })
    : [];
  const con = parciales.filter((v) => v != null);
  const promedio = con.length ? round1(con.reduce((a, b) => a + b, 0) / con.length) : null;
  const minimo = c.minimo != null ? Number(c.minimo) : escala * 0.6;
  let exento = false;
  let finalEsc = null;
  let valor = promedio;
  if (c.modo === 'final' && promedio != null) {
    exento = c.exentoDesde != null && c.exentoDesde !== '' && promedio >= Number(c.exentoDesde);
    const f = (c.final || {})[row];
    finalEsc = f != null && f !== '' ? round1((Number(f) * escala) / 100) : null;
    if (!exento) {
      const pp = Number(c.pesoParciales) || 0;
      const pf = Number(c.pesoFinal) || 0;
      valor = finalEsc == null || !(pp + pf) ? null : round1((promedio * pp + finalEsc * pf) / (pp + pf));
    }
  }
  let reportada = valor;
  if (valor != null && c.redondeo === 'entero') reportada = Math.round(valor);
  if (valor != null && c.reprobado5 && escala === 10 && valor < minimo) reportada = 5;
  return {
    parciales, promedio, final: finalEsc, exento, valor, reportada, minimo,
    aprobado: reportada == null ? null : reportada >= minimo,
    faltaFinal: c.modo === 'final' && promedio != null && !exento && finalEsc == null,
  };
}

export {
    STATUSES, READ, WORK, DEFAULT_CODES, TYPES, ASIST_ID, DEFAULT_SECTIONS, sectionsOf, newSectionId, secMin,
    activos, parseWorkbook, exportWorkbook, groupDates, sheetLayout, worksOf, normalizeWorks, normalizeGroup, hasData, discDates, datesOf, shortDate,
    DEFAULT_GRADING, gradeParts, finalGrade, gradingOn, weightOf, round1, parseHeader, parseDateHeader, cellText,
    studentReport, exportStudent, groupReport, exportGroupReport,
    enPeriodo, periodosListos, parcialDe, cierreVigente, SEMESTRAL, semestral,
};
