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
