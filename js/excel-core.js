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
 *   trab:    { '2026-11-14': { 5: 100, 6: 50 } }     // trabajo: 0 / 50 / 100
 * }
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
    { key: 'records', label: 'Asistencia', match: /ASIST/i },
    { key: 'lect', label: 'Lectura', match: /LECT/i },
    { key: 'trab', label: 'Trabajo', match: /TRAB/i },
  ];

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
    const cat = CATS.find((c) => c.match.test(m[4] || ''));
    return { date: `${y}-${pad(m[2])}-${pad(m[1])}`, cat: cat ? cat.key : 'records' };
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
      const data = { records: {}, lect: {}, trab: {} };
      const dateCols = [];
      const lastCol = Math.max(ws.columnCount, hdr.col);
      for (let c = hdr.col + 1; c <= lastCol; c++) {
        const h = parseHeader(ws.getCell(hdr.row, c).value);
        if (!h) continue;
        dateCols.push(c);
        const day = (data[h.cat][h.date] = data[h.cat][h.date] || {});
        for (const s of students) {
          const v = valueFromCell(ws.getCell(s.row, c).value, h.cat, codes);
          if (v != null) day[s.row] = v;
        }
      }

      // Quita días sin ningún dato (p. ej. la columna de lectura vacía de un día).
      for (const k of Object.keys(data)) {
        for (const d of Object.keys(data[k])) if (!Object.keys(data[k][d]).length) delete data[k][d];
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
        records: data.records, lect: data.lect, trab: data.trab,
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
    for (const c of CATS) {
      const m = g[c.key] || {};
      Object.keys(m).forEach((d) => { if (Object.keys(m[d]).length) set.add(d); });
    }
    return [...set].sort();
  }

  const hasData = (g, key, date) => Object.keys(((g[key] || {})[date]) || {}).length > 0;

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
    trDates.forEach((date, i) => out.push({ col: col++, date, cat: 'trab', first: i === 0, block: 'trab' }));
    return out;
  }

  // Escribe los datos en una copia del Excel original y devuelve el buffer .xlsx
  async function exportWorkbook(ExcelJS, originalBuffer, groups, codes = DEFAULT_CODES) {
    codes = Object.assign({}, DEFAULT_CODES, codes);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(originalBuffer);
    const thin = { style: 'thin' };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    for (const g of groups) {
      const ws = wb.getWorksheet(g.sheet);
      if (!ws) continue;

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

      columnLayout(g).forEach(({ col, date, cat }) => {
        const [y, m, d] = date.split('-');
        const label = CATS.find((c) => c.key === cat).label;
        const h = ws.getCell(g.headerRow, col);
        h.value = `${d}/${m}/${y.slice(-2)}\n${label}`;
        h.font = Object.assign({}, headerSrc.font, { bold: true });
        h.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        h.border = border;
        ws.getColumn(col).width = 11;

        const day = (g[cat] || {})[date] || {};
        for (const s of g.students) {
          const v = day[s.row];
          const cell = ws.getCell(s.row, col);
          if (v == null) cell.value = null;
          else cell.value = cat === 'trab' ? Number(v) : codeValue(v, codes);
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = border;
        }
      });
    }
    return wb.xlsx.writeBuffer();
  }

  return {
    STATUSES, READ, WORK, CATS, DEFAULT_CODES,
    parseWorkbook, exportWorkbook, groupDates, columnLayout, parseHeader, parseDateHeader, cellText,
  };
});
