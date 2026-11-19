// Qué sube la sincronización: solo lo que cambió, días borrados se mandan vacíos, nada sin id.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { normalizeGroup } from '../src/lib/excel-core.js';
import { faltaImportar, planSecciones, planDias, diasDeTrabajos } from '../src/lib/sync-plan.js';

function grupo() {
  const g = { sheet: 'V205', students: [{ row: 5, no: '1', name: 'ANA' }, { row: 6, no: '2', name: 'BETO' }], grading: { w: { records: 50, trab: 50 }, retardo: 100, escala: 10 } };
  normalizeGroup(g);
  g.records['2026-11-20'] = { 5: 'A', 6: 'F' };
  g.disc['2026-11-20'] = { 6: 0 };
  g.trab['2026-11-20'] = [{ title: 'Ensayo', vals: { 5: 100 } }, { title: '', vals: { 6: 50 } }];
  return g;
}
// Simula lo que hace la nube: asigna ids y guarda huellas.
function subir(g) {
  g.nube = g.nube || { h: {} };
  g.nube.clase_id = 'clase';
  g.students.forEach((s) => (s.alumno_id = `al-${s.row}`));
  const ps = planSecciones(g);
  if (ps) { g.sections.forEach((s) => (s.nube_id = `sec-${s.id}`)); g.nube.h.__secciones = ps.hash; }
  const p = planDias(g);
  p.olvidar.forEach((k) => delete g.nube.h[k]);
  p.trabajos.forEach((t) => t.works.forEach((w, i) => (w.nube_id = w.nube_id || `tr-${t.fecha}-${i}`)));
  [...p.dias, ...p.trabajos].forEach((x) => { if (x.hash == null) delete g.nube.h[x.key]; else g.nube.h[x.key] = x.hash; });
  return p;
}

test('la primera vez hay que importar y subir todo', () => {
  const g = grupo();
  assert.equal(faltaImportar(g), true);
  const p = subir(g);
  assert.equal(faltaImportar(g), false);
  assert.deepEqual(p.dias.map((d) => d.key).sort(), ['disc|2026-11-20', 'records|2026-11-20']);
  assert.deepEqual(p.dias.find((d) => d.key === 'records|2026-11-20').valores, { 'al-5': 'A', 'al-6': 'F' });
  assert.deepEqual(p.dias.find((d) => d.key === 'disc|2026-11-20').valores, { 'al-6': '0' });
  assert.equal(p.trabajos.length, 1);
  assert.deepEqual(diasDeTrabajos(g, p.trabajos[0]).map((d) => [d.trabajo, d.valores]),
    [['tr-2026-11-20-0', { 'al-5': '100' }], ['tr-2026-11-20-1', { 'al-6': '50' }]]);
});

test('después solo sube lo que cambió', () => {
  const g = grupo();
  subir(g);
  assert.equal(planSecciones(g), null);
  let p = planDias(g);
  assert.deepEqual([p.dias.length, p.trabajos.length], [0, 0]);
  g.records['2026-11-20'][6] = 'R';
  g.records['2026-11-21'] = { 5: 'A' };
  p = planDias(g);
  assert.deepEqual(p.dias.map((d) => d.key).sort(), ['records|2026-11-20', 'records|2026-11-21']);
  g.trab['2026-11-20'][0].title = 'Ensayo final';
  assert.equal(planDias(g).trabajos.length, 1);
});

test('un día borrado se manda vacío y una sección quitada se olvida', () => {
  const g = grupo();
  subir(g);
  delete g.records['2026-11-20'];
  g.sections = g.sections.filter((s) => s.id !== 'disc');
  delete g.disc;
  const p = planDias(g);
  assert.deepEqual(p.dias.map((d) => [d.key, d.valores, d.hash]), [['records|2026-11-20', {}, null]]);
  assert.deepEqual(p.olvidar, ['disc|2026-11-20']);
  assert.ok(planSecciones(g), 'las secciones cambiaron');
});

test('cambiar porcentaje o materia vuelve a mandar las secciones', () => {
  const g = grupo();
  subir(g);
  g.grading.w.records = 40;
  assert.ok(planSecciones(g));
  subir(g);
  g.subject = 'Ética';
  assert.equal(planSecciones(g).materia, 'Ética');
});

test('un alumno agregado en la app pide importar otra vez', () => {
  const g = grupo();
  subir(g);
  g.students.push({ row: 7, no: '3', name: 'CARLA', added: true });
  assert.equal(faltaImportar(g), true);
});

test('bajar de la nube arma el grupo igual y lo deja como ya subido', async () => {
  const { armarGrupo } = await import('../src/lib/sync-plan.js');
  // Un grupo recién leído del Excel (sin datos), en otro celular.
  const g = { sheet: 'V205', headerRow: 4, students: [{ row: 5, no: '1', name: 'ANA' }, { row: 6, no: '2', name: 'BETO' }] };
  normalizeGroup(g);
  g.records['2026-01-01'] = { 5: 'F' }; // dato viejo del Excel: la nube manda
  const nube = {
    clase: { id: 'c1', grupo_id: 'g1', materia: 'Ética', retardo: 50, escala: 100 },
    secciones: [
      { id: 's-asist', nombre: 'Asistencia', plantilla: 'asist', porcentaje: 40, minimo: null, orden: 0 },
      { id: 's-part', nombre: 'Participación', plantilla: 'sino', porcentaje: 20, minimo: 3, orden: 1 },
      { id: 's-tar', nombre: 'Tareas', plantilla: 'trab', porcentaje: 40, minimo: null, orden: 2 },
    ],
    inscripciones: [
      { alumno_id: 'a1', numero_lista: '1', nombre: 'Ana' },
      { alumno_id: 'a2', numero_lista: '2', nombre: 'Beto' },
      { alumno_id: 'a3', numero_lista: '3', nombre: 'Carla Nueva' },
    ],
    trabajos: [{ id: 't2', seccion_id: 's-tar', fecha: '2026-11-20', titulo: '', orden: 1 }, { id: 't1', seccion_id: 's-tar', fecha: '2026-11-20', titulo: 'Unit 1', orden: 0 }],
    capturas: [
      { seccion_id: 's-asist', alumno_id: 'a1', fecha: '2026-11-20', trabajo_id: null, valor: 'A' },
      { seccion_id: 's-asist', alumno_id: 'a3', fecha: '2026-11-20', trabajo_id: null, valor: 'R' },
      { seccion_id: 's-part', alumno_id: 'a2', fecha: '2026-11-20', trabajo_id: null, valor: 'S' },
      { seccion_id: 's-tar', alumno_id: 'a1', fecha: '2026-11-20', trabajo_id: 't1', valor: '100' },
      { seccion_id: 's-tar', alumno_id: 'a2', fecha: '2026-11-20', trabajo_id: 't2', valor: '50' },
    ],
  };
  nube.clase.config_semestral = { modo: 'final', pesoParciales: 60, pesoFinal: 40, final: { a2: 75 } };
  nube.cierres = [{ id: 'k1', tipo: 'parcial', parcial: 1, calificaciones: { a1: { final: 9, partes: { 's-asist': 100 } } }, criterios: { w: { 's-asist': 50, 's-tar': 50 }, retardo: 100, escala: 10 }, cerrado_en: '2026-10-01T00:00:00Z', reabierto_en: null }];
  armarGrupo(g, nube);
  assert.deepEqual(g.semestral.final, { 6: 75 });
  assert.deepEqual(g.cierres.map((c) => [c.tipo, c.parcial, c.calificaciones, c.nube_id]), [['parcial', 1, { 5: { final: 9, partes: { records: 100 } } }, 'k1']]);
  assert.deepEqual(g.cierres[0].criterios.w, { records: 50, [g.sections[2].id]: 50 });
  assert.equal(g.subject, 'Ética');
  assert.deepEqual(g.sections.map((s) => [s.name, s.type]), [['Asistencia', 'asist'], ['Participación', 'sino'], ['Tareas', 'trab']]);
  assert.equal(g.sections[1].min, 3);
  assert.deepEqual(g.grading, { w: { records: 40, [g.sections[1].id]: 20, [g.sections[2].id]: 40 }, retardo: 50, escala: 100 });
  assert.deepEqual(g.records, { '2026-11-20': { 5: 'A', 7: 'R' } });
  assert.deepEqual(g.students.map((s) => [s.row, s.name, s.alumno_id, !!s.added]), [[5, 'ANA', 'a1', false], [6, 'BETO', 'a2', false], [7, 'Carla Nueva', 'a3', true]]);
  const tar = g[g.sections[2].id]['2026-11-20'];
  assert.deepEqual(tar.map((w) => [w.title, w.vals, w.nube_id]), [['Unit 1', { 5: 100 }, 't1'], ['', { 6: 50 }, 't2']]);
  // No hay nada pendiente por subir.
  assert.equal(planSecciones(g), null);
  const p = planDias(g);
  assert.deepEqual([p.dias.length, p.trabajos.length, faltaImportar(g)], [0, 0, false]);
});

test('resumen de la clase para el director', async () => {
  const { resumenClase } = await import('../src/lib/sync-plan.js');
  const g = grupo();
  g.students.forEach((s) => (s.alumno_id = `al-${s.row}`));
  g.grading = { w: { records: 100 }, retardo: 100, escala: 10 };
  g.records['2026-11-21'] = { 5: 'A', 6: 'F' };
  const per = { semestre: 'S', parciales: [{ numero: 1, inicio: '2026-11-01', fin: '2026-11-30' }] };
  const { datos, ultima } = resumenClase(g, per);
  assert.equal(ultima, '2026-11-21');
  assert.deepEqual(datos.alumnos['al-5'], { f: 10, a: 100, F: 0, e: 'verde', r: [], p: [10], s: 10 });
  assert.equal(datos.alumnos['al-6'].e, 'rojo');
  assert.equal(datos.alumnos['al-6'].F, 2);
});

test('cambios de otro dispositivo: solo se baja si otro cambió después', async () => {
  const { hayQueBajar, ultimoCambio } = await import('../src/lib/sync-plan.js');
  const g = { nube: { clase_id: 'c1' } };
  const cambios = { yo: '2026-11-20T10:00:05.123456+00:00', compu: '2026-11-20T10:00:01.000001+00:00' };
  assert.equal(ultimoCambio(cambios), cambios.yo);
  assert.equal(ultimoCambio(cambios, 'yo'), cambios.compu);
  assert.equal(hayQueBajar(g, cambios, 'yo'), true); // nunca ha visto nada de la compu
  g.nube.visto = cambios.yo;
  assert.equal(hayQueBajar(g, cambios, 'yo'), false); // ya bajó todo hasta su último cambio
  assert.equal(hayQueBajar(g, { yo: cambios.yo }, 'yo'), false); // sus propios cambios no cuentan
  assert.equal(hayQueBajar(g, { ...cambios, compu: '2026-11-20T10:03:00+00:00' }, 'yo'), true);
  assert.equal(hayQueBajar({ nube: {} }, cambios, 'yo'), false); // grupo sin clase en la nube
});

test('actualizar un grupo ya ligado conserva sus secciones, alumnos y cuadernillo', async () => {
  const { armarGrupo } = await import('../src/lib/sync-plan.js');
  const g = grupo();
  subir(g);
  g.nube.pdf = 'marca-local';
  g.nube.pdf_nube = 'libro.pdf|100';
  const ids = g.sections.map((s) => s.id);
  // En la compu: BETO llegó con retardo, nuevo día y cambiaron el nombre de ANA en la lista.
  const nube = {
    clase: { id: 'clase', grupo_id: 'g1', materia: '', retardo: 100, escala: 10, cambios: { compu: '2026-11-21T08:00:00+00:00' } },
    secciones: g.sections.map((s, i) => ({ id: s.nube_id, nombre: s.name, plantilla: s.type, porcentaje: (g.grading.w[s.id] || 0), minimo: s.type === 'sino' ? 4 : null, orden: i })),
    inscripciones: [{ alumno_id: 'al-5', numero_lista: '1', nombre: 'ANA MARÍA' }, { alumno_id: 'al-6', numero_lista: '2', nombre: 'BETO', baja: '2026-11-22' }],
    trabajos: [],
    capturas: [
      { seccion_id: 'sec-records', alumno_id: 'al-5', fecha: '2026-11-20', trabajo_id: null, valor: 'A' },
      { seccion_id: 'sec-records', alumno_id: 'al-6', fecha: '2026-11-20', trabajo_id: null, valor: 'R' },
      { seccion_id: 'sec-records', alumno_id: 'al-6', fecha: '2026-11-21', trabajo_id: null, valor: 'F' },
    ],
  };
  armarGrupo(g, nube);
  assert.deepEqual(g.sections.map((s) => s.id), ids);
  assert.deepEqual(g.records, { '2026-11-20': { 5: 'A', 6: 'R' }, '2026-11-21': { 6: 'F' } });
  assert.deepEqual(g.students.map((s) => [s.row, s.alumno_id, !!s.added]), [[5, 'al-5', false], [6, 'al-6', false]]);
  assert.deepEqual([g.nube.pdf, g.nube.pdf_nube, g.nube.visto], ['marca-local', 'libro.pdf|100', '2026-11-21T08:00:00+00:00']);
  assert.equal(g.students[1].baja, '2026-11-22');
  assert.deepEqual(g.nube.bajas, { 'al-6': '2026-11-22' });
  assert.equal(planSecciones(g), null);
  const p = planDias(g);
  assert.deepEqual([p.dias.length, p.trabajos.length, faltaImportar(g)], [0, 0, false]);
});
