// Lista de ejemplo: dos grupos con alumnos inventados y unos días ya capturados, para conocer la app
// (pasar lista, calificar, resúmenes) sin el Excel propio. Se arma como un Excel de verdad para que
// exportar funcione igual. Queda marcada como ejemplo: no se sube a la cuenta.
import { loadExcel } from './util.js';
import { render } from './main-view.js';
import { save, state } from './state.js';

const NOMBRES = ['ANA SOFÍA', 'LUIS ÁNGEL', 'MARÍA JOSÉ', 'DIEGO', 'VALERIA', 'CARLOS ALBERTO', 'XIMENA', 'JUAN PABLO', 'CAMILA', 'MIGUEL',
  'DANIELA', 'FERNANDA', 'JESÚS ALONSO', 'ANDREA', 'EMILIANO', 'PAOLA', 'RICARDO', 'REGINA', 'ALEJANDRO', 'LUCÍA', 'SEBASTIÁN', 'RENATA', 'ITZEL', 'OMAR'];
const APELLIDOS = ['GARCÍA', 'LÓPEZ', 'MARTÍNEZ', 'VALENZUELA', 'RUIZ', 'SOTO', 'FÉLIX', 'ESPINOZA', 'MORALES', 'LEYVA', 'CASTRO', 'ROMERO',
  'ARMENTA', 'GÁMEZ', 'QUIJADA', 'BORBÓN', 'ACOSTA', 'MENDÍVIL', 'ZAZUETA', 'ORTIZ', 'VEGA', 'IBARRA'];

function azar(semilla) {
  let a = semilla;
  return () => { a = (a * 1664525 + 1013904223) % 4294967296; return a / 4294967296; };
}

// Días de clase anteriores a hoy (lunes a viernes).
function diasAnteriores(n) {
  const out = [];
  const d = new Date();
  while (out.length < n) {
    d.setDate(d.getDate() - 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
  }
  return out.reverse();
}

export async function listaDeEjemplo(cargarLibro) {
  const rnd = azar(7);
  const toma = (a) => a[Math.floor(rnd() * a.length)];
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  for (const [hoja, n] of [['1A', 24], ['3B', 22]]) {
    const ws = wb.addWorksheet(hoja);
    ws.addRow(['PREPARATORIA DE EJEMPLO']);
    ws.addRow([`GRUPO ${hoja} · LISTA DE EJEMPLO (alumnos inventados)`]);
    ws.addRow([]);
    ws.addRow(['NO.', 'NOMBRE DEL ALUMNO']);
    const nombres = new Set();
    while (nombres.size < n) nombres.add(`${toma(APELLIDOS)} ${toma(APELLIDOS)} ${toma(NOMBRES)}`);
    [...nombres].sort().forEach((nm, i) => ws.addRow([i + 1, nm]));
    ws.getColumn(2).width = 38;
  }
  const buf = await wb.xlsx.writeBuffer();
  if (!(await cargarLibro(buf, 'lista-de-ejemplo.xlsx', { ejemplo: true }))) return;

  // Unos días ya capturados para que los resúmenes y el semáforo tengan qué mostrar.
  const dias = diasAnteriores(6);
  for (const g of state.groups) {
    // Porcentajes de ejemplo (el maestro los cambia en Secciones y calificaciones).
    g.grading.w = { records: 20, lect: 20, trab: 40, disc: 10, exam: 10 };
    const nivel = Object.fromEntries(g.students.map((s) => [s.row, rnd() < 0.15 ? 0.55 : 0.92]));
    for (const d of dias) {
      g.records[d] = {};
      for (const s of g.students) {
        const x = rnd();
        g.records[d][s.row] = x < nivel[s.row] ? 'A' : x < nivel[s.row] + 0.05 ? 'R' : x < nivel[s.row] + 0.07 ? 'J' : 'F';
      }
    }
    const t = dias[dias.length - 2];
    g.trab[t] = [{ title: 'Resumen de lectura', vals: Object.fromEntries(g.students.map((s) => [s.row, nivel[s.row] > 0.6 ? (rnd() < 0.8 ? 100 : 50) : (rnd() < 0.5 ? 0 : 50)])) }];
    g.lect[dias[dias.length - 3]] = Object.fromEntries(g.students.map((s) => [s.row, rnd() < nivel[s.row] ? 'S' : 'N']));
  }
  save(true);
  render();
}
