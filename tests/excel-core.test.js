// Pruebas de la lógica de Excel y calificaciones (npm test). Usan una lista ficticia, nunca datos reales.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import * as L from '../src/lib/excel-core.js';

const NAMES = ['ACOSTA LUNA ANA', 'BELTRÁN RUIZ BRUNO', 'CRUZ SOTO CARLA', 'DÍAZ PAZ DIEGO', 'ESPINOZA MEZA EVA', 'FLORES GIL FÉLIX'];

// Excel como el del maestro: título, encabezado "No. / NOMBRE" en la fila 4 y alumnos debajo.
async function listaFicticia() {
  const wb = new ExcelJS.Workbook();
  for (const sheet of ['V205', 'V206']) {
    const ws = wb.addWorksheet(sheet);
    ws.getCell('A1').value = `GRUPO ${sheet} SEM B`;
    ws.getCell('A4').value = 'No.';
    ws.getCell('B4').value = 'NOMBRE';
    NAMES.forEach((n, i) => { ws.getCell(5 + i, 1).value = i + 1; ws.getCell(5 + i, 2).value = n; });
  }
  return wb.xlsx.writeBuffer();
}
const load = async (buf) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf); return wb; };
const W = { w: { records: 10, lect: 10, trab: 50, exam: 20, disc: 10 }, retardo: 100, escala: 100 };

test('lee los grupos y alumnos de la lista', async () => {
  const groups = L.parseWorkbook(await load(await listaFicticia()));
  assert.deepEqual(groups.map((g) => g.sheet), ['V205', 'V206']);
  assert.equal(groups[0].students.length, 6);
  assert.equal(groups[0].students[1].name, 'BELTRÁN RUIZ BRUNO');
  assert.equal(groups[0].headerRow, 4);
  assert.equal(groups[0].startCol, 3);
});

test('exportar y volver a importar conserva los datos (secciones de siempre)', async () => {
  const buf = await listaFicticia();
  const g = L.parseWorkbook(await load(buf))[0];
  L.normalizeGroup(g);
  const S = g.students;
  ['2026-11-10', '2026-11-11'].forEach((d, i) => {
    g.records[d] = {}; g.lect[d] = {};
    S.forEach((s, j) => { g.records[d][s.row] = 'AFRJ'[(i + j) % 4]; if (j % 2) g.lect[d][s.row] = 'S'; });
  });
  g.trab['2026-11-10'] = [{ title: 'Ensayo', vals: { [S[0].row]: 100, [S[1].row]: 50 } }, { title: '', vals: { [S[0].row]: 0 } }];
  g.disc['2026-11-11'] = { [S[2].row]: 0 };
  g.exam['2026-11-11'] = { [S[0].row]: 85, [S[3].row]: 60 };
  g.grading = W;
  const out = await L.exportWorkbook(ExcelJS, buf, [g], L.DEFAULT_CODES);
  const g2 = L.parseWorkbook(await load(out), L.DEFAULT_CODES, () => g.sections)[0];
  L.normalizeGroup(g2);
  for (const k of ['records', 'lect', 'disc', 'exam']) assert.deepEqual(g2[k], g[k], k);
  assert.deepEqual(g2.trab['2026-11-10'].map((w) => w.title), ['Ensayo', '']);
  assert.equal(g2.dateCols.length, L.sheetLayout(g).length);
});

test('secciones propias (inglés: examen oral y escrito) se exportan y se reconocen', async () => {
  const buf = await listaFicticia();
  const e = L.parseWorkbook(await load(buf))[0];
  e.sections = [
    { id: 'records', name: 'Asistencia', type: 'asist' },
    { id: 's_part', name: 'Participación', type: 'sino', min: 5 },
    { id: 's_tar', name: 'Tareas', type: 'trab' },
    { id: 's_oral', name: 'Examen oral', type: 'num' },
    { id: 's_esc', name: 'Examen escrito', type: 'num' },
  ];
  L.normalizeGroup(e);
  const [a, b] = e.students;
  e.records['2026-11-10'] = { [a.row]: 'A', [b.row]: 'A' };
  e.s_part['2026-11-10'] = { [a.row]: 'S', [b.row]: 'N' };
  e.s_tar['2026-11-11'] = [{ title: 'Unit 1', vals: { [a.row]: 100, [b.row]: 50 } }];
  e.s_oral['2026-11-12'] = { [a.row]: 90, [b.row]: 60 };
  e.s_esc['2026-11-13'] = { [a.row]: 80, [b.row]: 40 };
  e.grading = { w: { records: 10, s_part: 10, s_tar: 20, s_oral: 30, s_esc: 30 }, retardo: 100, escala: 10 };
  const out = await L.exportWorkbook(ExcelJS, buf, [e], L.DEFAULT_CODES);
  const ws = (await load(out)).getWorksheet('V205');
  const headers = L.sheetLayout(e).map((x) => String(ws.getCell(e.headerRow, x.col).value).replace(/\n/g, ' '));
  assert.ok(headers.includes('12/11/26 Examen oral'));
  assert.ok(headers.includes('Examen escrito 30%'));
  const e2 = L.parseWorkbook(await load(out), L.DEFAULT_CODES, () => e.sections)[0];
  L.normalizeGroup(e2);
  for (const k of ['records', 's_part', 's_tar', 's_oral', 's_esc']) assert.deepEqual(e2[k], e[k], k);
  // Participación: 1 "Sí" de mínimo 5 = 20 %.
  assert.equal(L.finalGrade(e, a.row, e.grading).parts.s_part, 20);
});

test('calificación: empieza completa y solo descuenta lo perdido', () => {
  const g = { sheet: 'X', students: [{ row: 5, no: '1', name: 'A' }, { row: 6, no: '2', name: 'B' }] };
  L.normalizeGroup(g);
  // Participaciones a la mitad (2 de mínimo 4), sin trabajos ni examen: 100 − 5 = 95.
  g.records['2026-11-16'] = { 5: 'A' }; g.records['2026-11-17'] = { 5: 'A' };
  g.lect['2026-11-16'] = { 5: 'S' }; g.lect['2026-11-17'] = { 5: 'S' };
  g.lect['2026-11-18'] = { 5: 'N' }; g.lect['2026-11-19'] = { 5: 'N' };
  assert.equal(L.finalGrade(g, 5, W).final, 95);
  assert.equal(L.finalGrade(g, 5, { ...W, escala: 10 }).final, 9.5);
  // Con todo capturado es el promedio ponderado normal.
  g.trab['2026-11-16'] = [{ title: '', vals: { 5: 100 } }]; g.trab['2026-11-17'] = [{ title: '', vals: { 5: 50 } }];
  g.exam['2026-11-17'] = { 5: 70 }; g.disc['2026-11-17'] = { 5: 0 };
  const r = L.finalGrade(g, 5, W);
  const p = r.parts;
  assert.equal(r.final, Math.round((10 * p.records + 10 * p.lect + 50 * p.trab + 20 * p.exam + 10 * p.disc) / 10) / 10);
  // Sin ningún dato o sin porcentajes: sin calificación.
  assert.equal(L.finalGrade(g, 6, W).final, null);
  assert.equal(L.finalGrade(g, 5, { w: {} }).final, null);
});

test('un solo examen en 0 que vale 20 % deja 80', () => {
  const g = { sheet: 'X', students: [{ row: 5, no: '1', name: 'A' }] };
  L.normalizeGroup(g);
  g.exam['2026-11-16'] = { 5: 0 };
  assert.equal(L.finalGrade(g, 5, W).final, 80);
});

test('resumen por alumno y del grupo generan su Excel', async () => {
  const g = L.parseWorkbook(await load(await listaFicticia()))[0];
  L.normalizeGroup(g);
  g.subject = 'Literatura';
  g.records['2026-11-10'] = Object.fromEntries(g.students.map((s, j) => [s.row, j % 3 ? 'A' : 'F']));
  g.grading = W;
  const rep = L.groupReport(g, W);
  assert.equal(rep.stats.alumnos, 6);
  assert.equal(rep.stats.faltas, 2);
  const s = await L.exportStudent(ExcelJS, g, g.students[0].row, W);
  const ws = (await load(s)).worksheets[0];
  assert.equal(ws.getCell('B6').value, 'Literatura');
  assert.ok((await L.exportGroupReport(ExcelJS, g, W)).byteLength > 1000);
});

test('parciales y calificación semestral', () => {
  const g = { sheet: 'X', students: [{ row: 5, no: '1', name: 'A' }, { row: 6, no: '2', name: 'B' }] };
  L.normalizeGroup(g);
  const per = { semestre: 'Ago–Dic 2026', parciales: [
    { numero: 1, inicio: '2026-08-17', fin: '2026-09-30' },
    { numero: 2, inicio: '2026-10-01', fin: '2026-11-10' },
    { numero: 3, inicio: '2026-11-11', fin: '2026-12-18' },
  ] };
  const gr = { w: { exam: 100 }, retardo: 100, escala: 10 };
  g.exam['2026-09-10'] = { 5: 80, 6: 50 };
  g.exam['2026-10-20'] = { 5: 90, 6: 60 };
  g.exam['2026-12-01'] = { 5: 100, 6: 70 };
  assert.equal(L.parcialDe(per, '2026-10-20'), 2);
  assert.equal(L.finalGrade(L.enPeriodo(g, '2026-10-01', '2026-11-10'), 5, gr).final, 9);
  // Promedio de parciales.
  let r = L.semestral(g, 5, gr, per, {});
  assert.deepEqual([r.parciales, r.promedio, r.reportada, r.aprobado], [[8, 9, 10], 9, 9, true]);
  r = L.semestral(g, 6, gr, per, {});
  assert.deepEqual([r.promedio, r.aprobado], [6, true]);
  // Parciales 60 % + examen final 40 %, exento con 8.5 o más.
  const cfg = { modo: 'final', pesoParciales: 60, pesoFinal: 40, exentoDesde: 8.5, final: { 6: 40 } };
  r = L.semestral(g, 5, gr, per, cfg);
  assert.deepEqual([r.exento, r.reportada], [true, 9]);
  r = L.semestral(g, 6, gr, per, cfg);
  assert.deepEqual([r.exento, r.final, r.valor, r.aprobado], [false, 4, 5.2, false]);
  // Redondeo a entero y reprobados como 5.
  r = L.semestral(g, 6, gr, per, { ...cfg, redondeo: 'entero', reprobado5: true, final: { 6: 55 } });
  assert.deepEqual([r.valor, r.reportada], [5.8, 5]);
  // Sin examen final capturado: falta.
  r = L.semestral(g, 6, gr, per, { ...cfg, final: {} });
  assert.deepEqual([r.faltaFinal, r.reportada], [true, null]);
});

test('el Excel lleva parciales y semestral, y se reconocen al volver a subirlo', async () => {
  const buf = await listaFicticia();
  const g = L.parseWorkbook(await load(buf))[0];
  L.normalizeGroup(g);
  g.grading = { w: { exam: 100 }, retardo: 100, escala: 10 };
  const per = { semestre: 'S', parciales: [{ numero: 1, inicio: '2026-08-01', fin: '2026-09-30' }, { numero: 2, inicio: '2026-10-01', fin: '2026-12-31' }] };
  const [a] = g.students;
  g.exam['2026-09-10'] = { [a.row]: 80 };
  g.exam['2026-10-10'] = { [a.row]: 100 };
  g.cierres = [{ tipo: 'parcial', parcial: 1, calificaciones: { [a.row]: { final: 7 } }, reabierto_en: null }];
  const out = await L.exportWorkbook(ExcelJS, buf, [g], L.DEFAULT_CODES, null, per);
  const ws = (await load(out)).getWorksheet('V205');
  const lay = L.sheetLayout(g, g.grading, per);
  const col = (k) => lay.find((x) => x.key === k).col;
  assert.equal(ws.getCell(g.headerRow, col('p1')).value, 'Parcial 1');
  assert.equal(ws.getCell(a.row, col('p1')).value, 7, 'parcial 1 cerrado manda');
  assert.equal(ws.getCell(a.row, col('p2')).value, 10);
  assert.equal(ws.getCell(a.row, col('semestral')).value, 8.5);
  const g2 = L.parseWorkbook(await load(out), L.DEFAULT_CODES, () => g.sections)[0];
  assert.equal(g2.dateCols.length, lay.length);
});

test('bajas: el alumno dado de baja sale de la lista desde su fecha y del resumen', async () => {
  const { activos, groupReport, normalizeGroup } = await import('../src/lib/excel-core.js');
  const g = { sheet: 'X', headerRow: 1, students: [{ row: 2, no: '1', name: 'UNO' }, { row: 3, no: '2', name: 'DOS', baja: '2026-11-25' }] };
  normalizeGroup(g);
  assert.deepEqual(activos(g, '2026-11-24').map((s) => s.name), ['UNO', 'DOS']);
  assert.deepEqual(activos(g, '2026-11-25').map((s) => s.name), ['UNO']);
  assert.deepEqual(activos(g).map((s) => s.name), ['UNO']);
  assert.equal(groupReport(g, {}).rows.length, 1);
});

test('lista nueva del mismo grupo: conserva lo capturado, cierres y examen final; los que se van quedan de baja', () => {
  const st = (row, no, name, extra = {}) => ({ row, no, name, ...extra });
  const prev = {
    sheet: '1A', headerRow: 4,
    students: [st(5, '1', 'ANA LÓPEZ', { alumno_id: 'a1' }), st(6, '2', 'BETO RUIZ', { alumno_id: 'a2' }), st(7, '3', 'CARLA SOTO', { alumno_id: 'a3' })],
    records: { '2026-11-10': { 5: 'A', 6: 'F', 7: 'R' } },
    lect: { '2026-11-11': { 5: 'S', 7: 'N' } },
    trab: { '2026-11-12': [{ title: 'Ensayo', vals: { 5: 100, 6: 50 }, nube_id: 't1' }] },
    cierres: [{ tipo: 'parcial', parcial: 1, calificaciones: { 5: { final: 9 }, 6: { final: 6 }, 7: { final: 8 } } }],
    semestral: { modo: 'final', final: { 6: 70 } },
  };
  // Lista nueva: cambia el orden, se va Carla y llega Diana (y Beto viene sin acento distinto).
  const g = { sheet: '1A', headerRow: 4, students: [st(5, '1', 'BETO  RUIZ'), st(6, '2', 'ANA LOPEZ'), st(7, '3', 'DIANA VEGA')], sections: L.DEFAULT_SECTIONS() };
  L.normalizeGroup(g);
  const r = L.actualizarGrupo(prev, g, '2026-11-20');
  assert.deepEqual(r, { llegan: ['DIANA VEGA'], seVan: ['CARLA SOTO'], siguen: ['ANA LOPEZ', 'BETO  RUIZ'] });
  const fila = (n) => g.students.find((s) => s.name.startsWith(n)).row;
  assert.equal(g.students.find((s) => s.name === 'ANA LOPEZ').alumno_id, 'a1');
  const carla = g.students.find((s) => s.name === 'CARLA SOTO');
  assert.deepEqual([carla.baja, carla.added, carla.alumno_id], ['2026-11-20', true, 'a3']);
  assert.deepEqual(g.records['2026-11-10'], { [fila('ANA')]: 'A', [fila('BETO')]: 'F', [fila('CARLA')]: 'R' });
  assert.deepEqual(g.lect['2026-11-11'], { [fila('ANA')]: 'S', [fila('CARLA')]: 'N' });
  assert.deepEqual(g.trab['2026-11-12'][0].vals, { [fila('ANA')]: 100, [fila('BETO')]: 50 });
  assert.equal(g.trab['2026-11-12'][0].nube_id, 't1');
  assert.equal(g.cierres[0].calificaciones[fila('BETO')].final, 6);
  assert.deepEqual(g.semestral.final, { [fila('BETO')]: 70 });
  // Diana no trae nada y sí aparece al pasar lista; Carla ya no.
  assert.deepEqual(L.activos(g).map((s) => s.name), ['BETO  RUIZ', 'ANA LOPEZ', 'DIANA VEGA']);
});
