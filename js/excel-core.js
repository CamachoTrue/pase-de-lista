/*
 * Lógica de Excel: leer la lista del maestro y escribir las columnas de asistencia.
 * Funciona en el navegador (window.ListaExcel) y en Node (module.exports) para pruebas.
 *
 * Modelo de un grupo:
 * {
 *   sheet: 'V205', title: 'GRUPO V205 SEM B 20-21',
 *   headerRow: 4, nameCol: 2, startCol: 3, dateCols: [..],
 *   students: [{ row: 5, no: '1', name: 'ACOSTA SANCHEZ ...' }],
 *   records: { '2026-11-14': { 5: 'A', 6: 'F' } }   // fila -> estado
 * }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ListaExcel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const STATUSES = ['A', 'F', 'R', 'J'];
  const DEFAULT_CODES = { A: 1, F: 0, R: 2, J: 3 };

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

  // Devuelve 'YYYY-MM-DD' si el encabezado es una fecha (celda fecha o texto dd/mm/aaaa).
  function parseDateHeader(v) {
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
    }
    if (v && typeof v === 'object' && v.result instanceof Date) return parseDateHeader(v.result);
    const m = clean(cellText(v)).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2}|\d{4})$/);
    if (!m) return null;
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    return `${y}-${pad(m[2])}-${pad(m[1])}`;
  }

  function codeToStatus(value, codes) {
    const t = clean(cellText(value)).toUpperCase();
    if (t === '') return null;
    for (const s of STATUSES) {
      if (String(codes[s]).toUpperCase() === t) return s;
    }
    return null;
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

      // Columnas de fecha ya existentes (por si suben un Excel exportado por la app).
      const records = {};
      const dateCols = [];
      const lastCol = Math.max(ws.columnCount, hdr.col);
      for (let c = hdr.col + 1; c <= lastCol; c++) {
        const date = parseDateHeader(ws.getCell(hdr.row, c).value);
        if (!date) continue;
        dateCols.push(c);
        const day = (records[date] = records[date] || {});
        for (const s of students) {
          const st = codeToStatus(ws.getCell(s.row, c).value, codes);
          if (st) day[s.row] = st;
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

      groups.push({ sheet: ws.name, title, headerRow: hdr.row, nameCol: hdr.col, startCol, dateCols, students, records });
    });
    return groups;
  }

  function codeValue(status, codes) {
    const v = codes[status];
    if (v === '' || v == null) return null;
    const n = Number(v);
    return Number.isFinite(n) && String(v).trim() !== '' ? n : String(v);
  }

  // Escribe la asistencia en una copia del Excel original y devuelve el buffer .xlsx
  async function exportWorkbook(ExcelJS, originalBuffer, groups, codes = DEFAULT_CODES) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(originalBuffer);
    const thin = { style: 'thin' };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    for (const g of groups) {
      const ws = wb.getWorksheet(g.sheet);
      if (!ws) continue;

      // Limpia columnas de fecha que traía el archivo original (se reescriben ordenadas).
      for (const c of g.dateCols) {
        ws.getCell(g.headerRow, c).value = null;
        for (const s of g.students) ws.getCell(s.row, c).value = null;
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

      const dates = Object.keys(g.records)
        .filter((d) => Object.keys(g.records[d]).length > 0)
        .sort();
      const headerSrc = ws.getCell(g.headerRow, g.nameCol);

      dates.forEach((date, i) => {
        const col = g.startCol + i;
        const [y, m, d] = date.split('-').map(Number);
        const h = ws.getCell(g.headerRow, col);
        h.value = new Date(Date.UTC(y, m - 1, d));
        h.numFmt = 'dd/mm/yy';
        h.font = Object.assign({}, headerSrc.font, { bold: true });
        h.alignment = { horizontal: 'center', vertical: 'middle' };
        h.border = border;
        ws.getColumn(col).width = 9.5;

        for (const s of g.students) {
          const st = g.records[date][s.row];
          const cell = ws.getCell(s.row, col);
          cell.value = st ? codeValue(st, codes) : null;
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = border;
        }
      });
    }
    return wb.xlsx.writeBuffer();
  }

  return { STATUSES, DEFAULT_CODES, parseWorkbook, exportWorkbook, parseDateHeader, cellText };
});
