/*
 * Lógica de Excel: leer la lista del maestro y escribir las columnas de cada día.
 * Funciona en el navegador (window.ListaExcel) y en Node (module.exports) para pruebas.
 *
 * Columnas por fecha: primero Asistencia + Lectura juntas (2 por día); después,
 * separadas por una columna vacía, las de Trabajo (1 por día).
 *
 * Modelo de un grupo:
 * {
 *   sheet: 'V205', title: 'GRUPO V205 SEM B 20-21',
 *   headerRow: 4, nameCol: 2, startCol: 3, dateCols: [..],
 *   students: [{ row: 5, no: '1', name: 'ACOSTA SANCHEZ ...', added?: true }],
 *   records: { '2026-11-14': { 5: 'A', 6: 'F' } },   // asistencia: A/F/R/J
 *   lect:    { '2026-11-14': { 5: 'S', 6: 'N' } },   // lectura: S (sí) / N (no)
 *   trab:    { '2026-11-14': [                        // trabajos del día (puede haber varios)
 *              { title: 'Ensayo', vals: { 5: 100, 6: 50 } },   // 0 / 50 / 100
 *            ] },
 *   disc:    { '2026-11-14': { 6: 0 } },             // disciplina: solo se guardan los 0 (los demás tienen 10)
 *   exam:    { '2026-11-14': { 5: 60, 6: 100 } }     // examen: porcentaje 0-100
 * }
 *
 * Disciplina no lleva una columna por día: al final van 2 columnas de resumen,
 * "Días con 0" (cantidad) y "Fechas" (las fechas en que sacó 0).
 * Examen tampoco: solo hay columna en las fechas en que hubo examen.
 * Al final de todo, si el maestro definió porcentajes, van las calificaciones (por apartado y total).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ListaExcel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const STATUSES = ['A', 'F', 'R', 'J'];
  const READ = ['S', 'N'];
  const WORK = [0, 50, 100];
  const DEFAULT_CODES = { A: 1, F: 0, R: 2, J: 3, S: 1, N: 0 };
  // Categorías por día, en el orden de las columnas del Excel.
  const CATS = [
    { key: 'records', label: 'Asistencia', match: /^ASIST/i },
    { key: 'lect', label: 'Lectura', match: /^LECT/i },
    { key: 'trab', label: 'Trabajo', match: /^TRAB/i },
  ];

  // Apartados que no van por día; sus encabezados también llevan fecha.
  const EXAM = { key: 'exam', label: 'Examen', match: /^EXAM/i };

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
  const pad = (n) => String(n).padStart(2, '0');
  const isoOf = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;

  // Encabezado de una columna de día: celda fecha (versión anterior = asistencia)
  // o texto "dd/mm/aa" + "Asistencia" | "Lectura" | "Trabajo".
  function parseHeader(v) {
    if (v instanceof Date && !isNaN(v)) return { date: isoOf(v), cat: 'records' };
    if (v && typeof v === 'object' && v.result instanceof Date) return parseHeader(v.result);
    const t = clean(cellText(v));
    const m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2}|\d{4})\b\s*(.*)$/);
    if (!m) return null;
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    const cat = [...CATS, EXAM].find((c) => c.match.test(m[4] || ''));
    const out = { date: `${y}-${pad(m[2])}-${pad(m[1])}`, cat: cat ? cat.key : 'records' };
    // Título del trabajo: "Trabajo: Ensayo" o "Trabajo 2: Ensayo" -> "Ensayo".
    if (out.cat === 'trab') out.title = (m[4] || '').replace(/^TRABAJO\s*\d*\s*:?\s*/i, '').trim();
    return out;
  }

  // ---------- Trabajos: varios por día, cada uno con título ----------
  // g.trab[fecha] = [{ title, vals: { fila: 0|50|100 } }]
  const worksOf = (g, date) => {
    const w = (g.trab || {})[date];
    if (Array.isArray(w)) return w;
    return w && typeof w === 'object' ? [{ title: '', vals: w }] : []; // formato anterior: un trabajo por día
  };
  const workUsed = (w) => !!w && (Object.keys(w.vals || {}).length > 0 || !!(w.title || '').trim());
  // Hay datos de un apartado en una fecha (para trabajos: algún trabajo con calificaciones).
  function hasData(g, key, date) {
    if (key === 'trab') return worksOf(g, date).some((w) => Object.keys(w.vals || {}).length);
    return Object.keys(((g[key] || {})[date]) || {}).length > 0;
  }
  // Pasa datos de la versión anterior ({fecha: {fila: v}}) al formato de varios trabajos.
  function normalizeWorks(g) {
    const t = g.trab || {};
    for (const d of Object.keys(t)) if (!Array.isArray(t[d])) t[d] = [{ title: '', vals: t[d] || {} }];
    g.trab = t;
    return g;
  }
  // Compatibilidad con pruebas/llamadas anteriores.
  const parseDateHeader = (v) => { const h = parseHeader(v); return h ? h.date : null; };

  function valueFromCell(value, cat, codes) {
    const t = clean(cellText(value)).toUpperCase();
    if (t === '') return null;
    if (cat === 'trab') {
      const n = Number(t);
      return WORK.includes(n) ? n : null;
    }
    if (cat === 'exam') {
      const n = Number(t);
      return Number.isFinite(n) && n >= 0 && n <= 100 ? n : null;
    }
    const keys = cat === 'lect' ? READ : STATUSES;
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

  function parseWorkbook(wb, codes = DEFAULT_CODES) {
    codes = Object.assign({}, DEFAULT_CODES, codes);
    const groups = [];
    wb.eachSheet((ws) => {
      const hdr = findHeader(ws);
      if (!hdr) return;

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

      // Columnas de día ya existentes (por si suben un Excel exportado por la app).
      const data = { records: {}, lect: {}, trab: {}, disc: {}, exam: {} };
      const dateCols = [];
      const lastCol = Math.max(ws.columnCount, hdr.col);
      for (let c = hdr.col + 1; c <= lastCol; c++) {
        // Resumen de disciplina: de la columna "Fechas" se recuperan los días con 0.
        const ht = clean(cellText(ws.getCell(hdr.row, c).value));
        // Columnas de calificación escritas por la app: se reconocen para poder reescribirlas.
        if (/^CALIFICACI[OÓ]N SOBRE \d+$/i.test(ht) || /^(ASISTENCIA|LECTURAS|TRABAJOS|DISCIPLINA|EXAMEN) \d+(\.\d+)?%$/i.test(ht)) {
          dateCols.push(c);
          continue;
        }
        if (/^DISCIPLINA (D[IÍ]AS CON 0|FECHAS)$/i.test(ht)) {
          dateCols.push(c);
          if (/FECHA/i.test(ht)) {
            for (const s of students) {
              const t = cellText(ws.getCell(s.row, c).value);
              for (const m of t.matchAll(/(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})/g)) {
                const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
                const d = `${y}-${pad(m[2])}-${pad(m[1])}`;
                (data.disc[d] = data.disc[d] || {})[s.row] = 0;
              }
            }
          }
          continue;
        }
        const h = parseHeader(ws.getCell(hdr.row, c).value);
        if (!h) continue;
        dateCols.push(c);
        let day;
        if (h.cat === 'trab') {
          // Cada columna de trabajo es un trabajo distinto de esa fecha.
          const w = { title: h.title || '', vals: {} };
          (data.trab[h.date] = data.trab[h.date] || []).push(w);
          day = w.vals;
        } else day = (data[h.cat][h.date] = data[h.cat][h.date] || {});
        for (const s of students) {
          const v = valueFromCell(ws.getCell(s.row, c).value, h.cat, codes);
          if (v != null) day[s.row] = v;
        }
      }

      // Quita días sin ningún dato (p. ej. la columna de lectura vacía de un día).
      for (const k of Object.keys(data)) {
        for (const d of Object.keys(data[k])) {
          if (k === 'trab') {
            data.trab[d] = data.trab[d].filter(workUsed);
            if (!data.trab[d].length) delete data.trab[d];
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

      groups.push({
        sheet: ws.name, title, headerRow: hdr.row, nameCol: hdr.col, startCol, dateCols, students,
        records: data.records, lect: data.lect, trab: data.trab, disc: data.disc, exam: data.exam,
      });
    });
    return groups;
  }

  function codeValue(key, codes) {
    const v = codes[key];
    if (v === '' || v == null) return null;
    const n = Number(v);
    return Number.isFinite(n) && String(v).trim() !== '' ? n : String(v);
  }

  // Fechas con al menos un dato en cualquiera de las 3 categorías, ordenadas.
  function groupDates(g) {
    const set = new Set();
    for (const c of CATS) Object.keys(g[c.key] || {}).forEach((d) => { if (hasData(g, c.key, d)) set.add(d); });
    return [...set].sort();
  }



  // Columnas que se escriben en el Excel, en orden: [{ col, date, cat, first, block }].
  // block 'al' = Asistencia + Lectura; block 'trab' = Trabajos (a la derecha, tras una columna vacía).
  function columnLayout(g) {
    const all = groupDates(g);
    const alDates = all.filter((d) => hasData(g, 'records', d) || hasData(g, 'lect', d));
    const trDates = all.filter((d) => hasData(g, 'trab', d));
    const out = [];
    let col = g.startCol;
    alDates.forEach((date) => {
      out.push({ col: col++, date, cat: 'records', first: true, block: 'al' });
      out.push({ col: col++, date, cat: 'lect', first: false, block: 'al' });
    });
    if (alDates.length && trDates.length) col++; // columna vacía de separación
    // Una columna por trabajo (con calificaciones) de cada fecha.
    let firstTr = true;
    trDates.forEach((date) => worksOf(g, date).forEach((w, work) => {
      if (!Object.keys(w.vals || {}).length) return;
      out.push({ col: col++, date, cat: 'trab', work, first: firstTr, block: 'trab' });
      firstTr = false;
    }));
    return out;
  }

  // Fechas (ordenadas) en que el alumno sacó 0 en disciplina.
  function discDates(g, row) {
    return Object.keys(g.disc || {}).filter((d) => g.disc[d][row] === 0).sort();
  }
  const shortDate = (iso) => { const [y, m, d] = iso.split('-'); return `${d}/${m}/${y.slice(-2)}`; };

  // Columnas del resumen de disciplina: al final, tras una columna vacía. null si no hay nada que escribir.
  const datesWith = (g, key) => Object.keys(g[key] || {}).filter((d) => Object.keys(g[key][d]).length).sort();

  function discCols(g) {
    const layout = columnLayout(g);
    if (!layout.length && !datesWith(g, 'disc').length && !datesWith(g, 'exam').length) return null;
    const count = layout.length ? layout[layout.length - 1].col + 2 : g.startCol;
    return { count, dates: count + 1 };
  }

  // Columnas de examen: una por fecha con examen, después del resumen de disciplina.
  function examCols(g) {
    const dc = discCols(g);
    if (!dc) return [];
    return datesWith(g, 'exam').map((date, i) => ({ col: dc.dates + 2 + i, date }));
  }

  // ---------- Calificaciones ----------
  // Apartados que se califican, en orden.
  const GRADE_PARTS = [
    { key: 'records', label: 'Asistencia' },
    { key: 'lect', label: 'Lecturas' },
    { key: 'trab', label: 'Trabajos' },
    { key: 'disc', label: 'Disciplina' },
    { key: 'exam', label: 'Examen' },
  ];
  const DEFAULT_GRADING = { w: { records: 0, lect: 0, trab: 0, disc: 0, exam: 0 }, retardo: 100, escala: 10, lectMin: 4 };
  const round1 = (n) => Math.round(n * 10) / 10;
  const avg = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : null);

  // Porcentaje (0-100) de cada apartado para un alumno; null si aún no hay datos.
  // Solo cuenta lo capturado: un día sin captura para ese alumno no suma ni resta.
  function gradeParts(g, row, grading) {
    grading = Object.assign({}, DEFAULT_GRADING, grading);
    const vals = (key) => Object.keys(g[key] || {}).map((d) => [d, g[key][d][row]]).filter(([, v]) => v != null);
    const att = vals('records');
    // Lecturas: con el mínimo de "Sí" que defina el maestro (o más) tiene el 100 %; debajo, proporcional.
    const lectVals = vals('lect');
    const lectMin = Math.max(1, Math.round(Number(grading.lectMin) || DEFAULT_GRADING.lectMin));
    const lectSi = lectVals.filter(([, v]) => v === 'S').length;
    // Asistencia y Justificado = 100 %, Retardo = lo que diga el maestro, Falta = 0 %.
    const attScore = { A: 100, J: 100, R: Number(grading.retardo) || 0, F: 0 };
    // Disciplina: días en que el alumno estuvo en clase (A o R) o tuvo reporte; 10 salvo los días con 0.
    const present = new Set(att.filter(([, v]) => v === 'A' || v === 'R').map(([d]) => d));
    const zeros = discDates(g, row);
    zeros.forEach((d) => present.add(d));
    return {
      records: avg(att.map(([, v]) => attScore[v])),
      lect: lectVals.length ? (100 * Math.min(lectSi, lectMin)) / lectMin : null,
      trab: avg(Object.keys(g.trab || {}).flatMap((d) => worksOf(g, d).map((w) => (w.vals || {})[row]))
        .filter((v) => v != null).map(Number)),
      disc: present.size ? (100 * (present.size - zeros.length)) / present.size : null,
      exam: avg(vals('exam').map(([, v]) => Number(v))),
    };
  }

  // Calificación final en la escala elegida (10 o 100). Los apartados sin datos aún no cuentan
  // y su porcentaje se reparte entre los demás, para que siempre haya una calificación al día.
  function finalGrade(g, row, grading) {
    grading = Object.assign({}, DEFAULT_GRADING, grading);
    const parts = gradeParts(g, row, grading);
    let sum = 0;
    let wsum = 0;
    for (const p of GRADE_PARTS) {
      const w = Number(grading.w[p.key]) || 0;
      if (w > 0 && parts[p.key] != null) { sum += w * parts[p.key]; wsum += w; }
    }
    const final = wsum ? round1(((sum / wsum) * (Number(grading.escala) || 10)) / 100) : null;
    return { parts, final };
  }
  const gradingOn = (grading) => !!grading && GRADE_PARTS.some((p) => Number((grading.w || {})[p.key]) > 0);

  // Columnas de calificación: una por apartado con porcentaje > 0 y la final. Van al final de todo.
  function gradeCols(g, grading) {
    if (!gradingOn(grading)) return [];
    const dc = discCols(g);
    const ex = examCols(g);
    let col = ex.length ? ex[ex.length - 1].col + 2 : dc ? dc.dates + 2 : g.startCol;
    const out = GRADE_PARTS.filter((p) => Number(grading.w[p.key]) > 0)
      .map((p) => ({ col: col++, key: p.key, label: `${p.label}\n${Number(grading.w[p.key])}%` }));
    out.push({ col: col++, key: 'final', label: `Calificación\nsobre ${Number(grading.escala) || 10}` });
    return out;
  }

  // Escribe los datos en una copia del Excel original y devuelve el buffer .xlsx
  async function exportWorkbook(ExcelJS, originalBuffer, groups, codes = DEFAULT_CODES, grading = null) {
    codes = Object.assign({}, DEFAULT_CODES, codes);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(originalBuffer);
    const thin = { style: 'thin' };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    for (const g of groups) {
      const ws = wb.getWorksheet(g.sheet);
      if (!ws) continue;
      normalizeWorks(g);

      // Limpia columnas de día que traía el archivo original (se reescriben ordenadas).
      for (const c of g.dateCols || []) {
        // Formato puesto por la app: se quita para que no queden bordes sueltos si cambia el acomodo.
        for (const r of [g.headerRow, ...g.students.map((s) => s.row)]) {
          const cell = ws.getCell(r, c);
          cell.value = null;
          cell.style = {};
        }
      }

      // Alumnos agregados en la app: se escriben al final de la lista con el mismo formato
      // (bordes, fuente) que el último alumno del archivo original.
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

      const dates = groupDates(g);
      const headerSrc = ws.getCell(g.headerRow, g.nameCol);
      if (dates.length) {
        const hr = ws.getRow(g.headerRow);
        hr.height = Math.max(hr.height || 15, 30); // dos renglones: fecha + categoría
      }

      // Línea gruesa al inicio de cada día (y al final de cada bloque); Lectura con gris claro.
      const thick = { style: 'medium', color: { argb: 'FF555555' } };
      const lectFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
      const layout = columnLayout(g);
      if (layout.some((x) => x.cat === 'trab' && (worksOf(g, x.date)[x.work].title || '').trim())) {
        const hr = ws.getRow(g.headerRow);
        hr.height = Math.max(hr.height || 15, 45); // fecha + "Trabajo: título" en varios renglones
      }
      layout.forEach(({ col, date, cat, block, work }, i) => {
        const next = layout[i + 1];
        const leftThick = cat === 'records' || (cat === 'trab' && layout[i].first);
        const rightThick = !next || next.block !== block;
        const cellBorder = Object.assign({}, border, leftThick ? { left: thick } : {}, rightThick ? { right: thick } : {});
        const [y, m, d] = date.split('-');
        let label = CATS.find((c) => c.key === cat).label;
        let wk = null;
        if (cat === 'trab') {
          wk = worksOf(g, date)[work];
          const n = worksOf(g, date).filter((w) => Object.keys(w.vals || {}).length).indexOf(wk) + 1;
          const title = (wk.title || '').trim();
          label = title ? `Trabajo: ${title}` : n > 1 ? `Trabajo ${n}` : 'Trabajo';
        }
        const h = ws.getCell(g.headerRow, col);
        h.value = `${d}/${m}/${y.slice(-2)}\n${label}`;
        // Estilo completo y propio de cada celda (ExcelJS comparte objetos de estilo entre celdas).
        h.style = {
          font: Object.assign({}, headerSrc.font, { bold: true }),
          alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
          border: cellBorder,
          ...(cat === 'lect' ? { fill: lectFill } : {}),
        };
        ws.getColumn(col).width = wk && (wk.title || '').trim() ? 16 : 11;

        const day = wk ? wk.vals : (g[cat] || {})[date] || {};
        for (const s of g.students) {
          const v = day[s.row];
          const cell = ws.getCell(s.row, col);
          if (v == null) cell.value = null;
          else cell.value = cat === 'trab' ? Number(v) : codeValue(v, codes);
          cell.style = {
            font: Object.assign({}, ws.getCell(s.row, g.nameCol).font || {}),
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: cellBorder,
            ...(cat === 'lect' ? { fill: lectFill } : {}),
          };
        }
      });

      // Resumen de disciplina.
      const dc = discCols(g);
      if (dc) {
        const thick = { style: 'medium', color: { argb: 'FF555555' } };
        [['count', 'Disciplina\nDías con 0'], ['dates', 'Disciplina\nFechas']].forEach(([k, label]) => {
          const col = dc[k];
          const b = Object.assign({}, border, k === 'count' ? { left: thick } : { right: thick });
          ws.getCell(g.headerRow, col).value = label;
          ws.getCell(g.headerRow, col).style = {
            font: Object.assign({}, headerSrc.font, { bold: true }),
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: b,
          };
          ws.getColumn(col).width = k === 'count' ? 11 : 24;
          for (const s of g.students) {
            const dates = discDates(g, s.row);
            const cell = ws.getCell(s.row, col);
            cell.value = k === 'count' ? dates.length : (dates.map(shortDate).join(', ') || null);
            cell.style = {
              font: Object.assign({}, ws.getCell(s.row, g.nameCol).font || {}),
              alignment: { horizontal: k === 'count' ? 'center' : 'left', vertical: 'middle', wrapText: true },
              border: b,
            };
          }
        });
      }

      // Exámenes.
      const ex = examCols(g);
      ex.forEach(({ col, date }, i) => {
        const thick = { style: 'medium', color: { argb: 'FF555555' } };
        const b = Object.assign({}, border, i === 0 ? { left: thick } : {}, i === ex.length - 1 ? { right: thick } : {});
        const h = ws.getCell(g.headerRow, col);
        h.value = `${shortDate(date)}\n${EXAM.label}`;
        h.style = {
          font: Object.assign({}, headerSrc.font, { bold: true }),
          alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
          border: b,
        };
        ws.getColumn(col).width = 11;
        for (const s of g.students) {
          const v = g.exam[date][s.row];
          const cell = ws.getCell(s.row, col);
          cell.value = v == null ? null : Number(v);
          cell.style = {
            font: Object.assign({}, ws.getCell(s.row, g.nameCol).font || {}),
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: b,
          };
        }
      });

      // Calificaciones (valores calculados por la app).
      // Cada grupo tiene sus propios porcentajes (g.grading); `grading` queda como respaldo.
      const gg = g.grading || grading;
      const gc = gradeCols(g, gg);
      gc.forEach(({ col, key, label }, i) => {
        const thick = { style: 'medium', color: { argb: 'FF555555' } };
        const isFinal = key === 'final';
        const b = Object.assign({}, border, i === 0 ? { left: thick } : {}, i === gc.length - 1 ? { right: thick } : {}, isFinal ? { left: thick } : {});
        const fill = isFinal ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } } : undefined;
        const h = ws.getCell(g.headerRow, col);
        h.value = label;
        h.style = {
          font: Object.assign({}, headerSrc.font, { bold: true }),
          alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
          border: b,
          ...(fill ? { fill } : {}),
        };
        ws.getColumn(col).width = isFinal ? 13 : 11;
        for (const s of g.students) {
          const r = finalGrade(g, s.row, gg);
          const v = isFinal ? r.final : r.parts[key];
          const cell = ws.getCell(s.row, col);
          cell.value = v == null ? null : round1(v);
          cell.style = {
            font: Object.assign({}, ws.getCell(s.row, g.nameCol).font || {}, isFinal ? { bold: true } : {}),
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: b,
            ...(fill ? { fill } : {}),
          };
        }
      });
    }
    return wb.xlsx.writeBuffer();
  }

  // ---------- Resumen de un alumno ----------
  // Todo lo capturado de un alumno, día por día, más su calificación. Lo usan la pantalla
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
    const student = g.students.find((s) => s.row === row) || null;
    const at = (key, d) => ((g[key] || {})[d] || {})[row];
    const dates = new Set();
    ['records', 'lect', 'disc', 'exam'].forEach((k) => Object.keys(g[k] || {}).forEach((d) => { if (at(k, d) != null) dates.add(d); }));
    Object.keys(g.trab || {}).forEach((d) => { if (worksOf(g, d).some((w) => (w.vals || {})[row] != null)) dates.add(d); });
    const counts = { A: 0, F: 0, R: 0, J: 0 };
    const lect = { S: 0, N: 0 };
    const works = [];
    const exams = [];
    const days = [...dates].sort().map((date) => {
      const att = at('records', date) || null;
      if (att) counts[att]++;
      const l = at('lect', date) || null;
      if (l) lect[l]++;
      const ws = worksOf(g, date);
      const dayWorks = ws.map((w, i) => ({ title: workLabel(w, i, ws.length), v: (w.vals || {})[row] }))
        .filter((w) => w.v != null).map((w) => ({ title: w.title, v: Number(w.v) }));
      dayWorks.forEach((w) => works.push({ date, ...w }));
      const exam = at('exam', date);
      if (exam != null) exams.push({ date, v: Number(exam) });
      // Disciplina: 0 si tuvo reporte; 10 si estuvo en clase (A o R) sin reporte; vacío si no vino.
      const disc = at('disc', date) === 0 ? 0 : att === 'A' || att === 'R' ? 10 : null;
      return { date, label: longDate(date), att, lect: l, works: dayWorks, disc, exam: exam != null ? Number(exam) : null };
    });
    const zeros = discDates(g, row);
    const { parts, final } = finalGrade(g, row, grading);
    const lectMin = Math.max(1, Math.round(Number(grading.lectMin) || DEFAULT_GRADING.lectMin));
    const detail = {
      records: `Asistencias ${counts.A} · Faltas ${counts.F} · Retardos ${counts.R} · Justificados ${counts.J}`,
      lect: lect.S + lect.N ? `Sí ${lect.S} de ${lect.S + lect.N} · mínimo para el 100 %: ${lectMin}` : 'Sin lecturas',
      trab: works.length ? `${works.length} trabajo(s) calificado(s)` : 'Sin trabajos',
      disc: zeros.length ? `${zeros.length} día(s) con 0: ${zeros.map(shortDate).join(', ')}` : 'Sin días con 0',
      exam: exams.length ? exams.map((e) => `${shortDate(e.date)}: ${e.v}`).join(' · ') : 'Sin exámenes',
    };
    return { student, days, counts, lect, works, exams, zeros, parts, final, detail, grading, graded: gradingOn(grading) };
  }

  // Excel nuevo (no es copia del original) con el concentrado de un solo alumno.
  function exportStudent(ExcelJS, g, row, grading) {
    const r = studentReport(g, row, grading);
    const s = r.student;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Resumen', { views: [{ showGridLines: false }] });
    ws.columns = [{ width: 18 }, { width: 14 }, { width: 12 }, { width: 34 }, { width: 12 }, { width: 10 }];
    const thin = { style: 'thin', color: { argb: 'FF999999' } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };
    const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
    const HEAD = fill('FFD9E2F3');
    const RED = fill('FFF8D7D7');
    const YEL = fill('FFFFF2CC');
    const style = (cell, o = {}) => {
      cell.style = {
        font: { name: 'Calibri', size: 11, bold: !!o.bold, color: o.color ? { argb: o.color } : undefined },
        alignment: { vertical: 'middle', horizontal: o.align || 'left', wrapText: true },
        ...(o.box ? { border } : {}),
        ...(o.fill ? { fill: o.fill } : {}),
      };
    };
    const put = (r0, c, v, o) => { const cell = ws.getCell(r0, c); cell.value = v; style(cell, o); return cell; };

    let y = 1;
    put(y, 1, 'Resumen del alumno', { bold: true });
    ws.getCell(y, 1).font = { name: 'Calibri', size: 16, bold: true };
    y += 2;
    const info = [['Alumno', s ? s.name : ''], ['No. de lista', s ? s.no : ''], ['Grupo', g.title ? `${g.sheet} — ${g.title}` : g.sheet], ...(g.subject ? [['Materia', g.subject]] : []),
      ['Generado', longDate((() => { const t = new Date(); return `${t.getFullYear()}-${t.getMonth() + 1}-${t.getDate()}`; })())]];
    info.forEach(([k, v]) => { put(y, 1, k, { bold: true }); ws.mergeCells(y, 2, y, 4); put(y, 2, v == null ? '' : String(v)); y++; });
    y++;

    // Calificación por apartado
    put(y, 1, 'Calificación', { bold: true });
    ws.getCell(y, 1).font = { name: 'Calibri', size: 13, bold: true };
    y++;
    ['Apartado', 'Vale', 'Resultado', 'Detalle'].forEach((t, i) => put(y, i + 1, t, { bold: true, box: true, fill: HEAD, align: i === 3 ? 'left' : 'center' }));
    y++;
    for (const p of GRADE_PARTS) {
      const w = Number(r.grading.w[p.key]) || 0;
      const v = r.parts[p.key];
      put(y, 1, p.label, { box: true, bold: true });
      put(y, 2, w ? `${w} %` : '—', { box: true, align: 'center' });
      put(y, 3, v == null ? 'Sin datos' : `${round1(v)} %`, { box: true, align: 'center' });
      put(y, 4, r.detail[p.key], { box: true });
      y++;
    }
    put(y, 1, `Calificación final (sobre ${r.grading.escala})`, { bold: true, box: true, fill: YEL });
    ws.mergeCells(y, 1, y, 2);
    put(y, 3, r.final == null ? (r.graded ? 'Sin datos' : 'Sin porcentajes') : r.final, { bold: true, box: true, fill: YEL, align: 'center' });
    put(y, 4, r.graded ? 'Los apartados sin datos aún no cuentan.' : 'El maestro aún no define cuánto vale cada apartado.', { box: true, fill: YEL });
    y += 2;

    // Detalle por día
    put(y, 1, 'Detalle por día', { bold: true });
    ws.getCell(y, 1).font = { name: 'Calibri', size: 13, bold: true };
    y++;
    ['Fecha', 'Asistencia', 'Lectura', 'Trabajos', 'Disciplina', 'Examen'].forEach((t, i) => put(y, i + 1, t, { bold: true, box: true, fill: HEAD, align: i === 3 ? 'left' : 'center' }));
    y++;
    if (!r.days.length) {
      ws.mergeCells(y, 1, y, 6);
      put(y, 1, 'Aún no hay datos capturados de este alumno.', { box: true });
    }
    for (const d of r.days) {
      put(y, 1, d.label, { box: true });
      put(y, 2, d.att ? STATUS_LABEL[d.att] : '', { box: true, align: 'center', fill: d.att === 'F' ? RED : undefined });
      put(y, 3, d.lect ? (d.lect === 'S' ? 'Sí' : 'No') : '', { box: true, align: 'center', fill: d.lect === 'N' ? RED : undefined });
      put(y, 4, d.works.map((w) => `${w.title}: ${w.v}`).join('\n'), { box: true, fill: d.works.some((w) => w.v === 0) ? RED : undefined });
      put(y, 5, d.disc, { box: true, align: 'center', fill: d.disc === 0 ? RED : undefined });
      put(y, 6, d.exam, { box: true, align: 'center' });
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
    const graded = gradingOn(grading);
    let present = 0;
    let taken = 0;
    const rows = g.students.map((s) => {
      const c = { A: 0, F: 0, R: 0, J: 0 };
      Object.values(g.records || {}).forEach((day) => { const v = day[s.row]; if (v in c) c[v]++; });
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
        alumnos: g.students.length,
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
    const parts = GRADE_PARTS.filter((p) => Number(r.grading.w[p.key]) > 0);
    const head = ['No.', 'Nombre', 'Asistencias', 'Faltas', 'Retardos', 'Justificados', '% Asistencia',
      ...parts.map((p) => `${p.label}\n${r.grading.w[p.key]}%`), `Calificación\nsobre ${r.escala}`, 'Estado', 'Motivo'];
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
        ...parts.map((p) => (row.parts[p.key] == null ? '' : round1(row.parts[p.key]))),
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

  return {
    STATUSES, READ, WORK, CATS, DEFAULT_CODES,
    parseWorkbook, exportWorkbook, groupDates, columnLayout, worksOf, normalizeWorks, hasData, discCols, discDates, examCols, shortDate,
    GRADE_PARTS, DEFAULT_GRADING, gradeParts, finalGrade, gradeCols, gradingOn, round1, parseHeader, parseDateHeader, cellText,
    studentReport, exportStudent, groupReport, exportGroupReport,
  };
});
