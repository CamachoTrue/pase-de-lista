// Qué falta subir de un grupo a la nube (lógica pura, con pruebas en tests/sync-plan.test.js).
// Por cada (sección, fecha) se guarda una huella de lo que ya se subió (g.nube.h); si lo de hoy es
// distinto, ese día se vuelve a mandar completo y la base de datos lo deja exactamente igual.
import { sectionsOf, worksOf, weightOf, secMin } from './excel-core.js';

export function huella(texto) {
  let h = 5381;
  for (let i = 0; i < texto.length; i++) h = ((h << 5) + h + texto.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36) + texto.length.toString(36);
}

// Valor tal como se guarda en la nube: A/F/R/J, S/N o número.
const valorNube = (type, v) => (type === 'asist' || type === 'sino' ? String(v) : String(Number(v)));

// { fila: valor } del celular → { alumno_id: valor } ordenado (los alumnos aún sin id se omiten).
function valores(g, type, dia) {
  const id = {};
  g.students.forEach((s) => { if (s.alumno_id) id[s.row] = s.alumno_id; });
  const out = {};
  Object.keys(dia || {})
    .filter((row) => id[row] && dia[row] != null)
    .sort((a, b) => (id[a] < id[b] ? -1 : 1))
    .forEach((row) => { out[id[row]] = valorNube(type, dia[row]); });
  return out;
}

// ¿Hay que crear la clase o dar de alta alumnos nuevos?
export const faltaImportar = (g) => !(g.nube && g.nube.clase_id) || g.students.some((s) => !s.alumno_id);

// Secciones y calificación de la clase. null si ya están al día.
export function planSecciones(g) {
  const h = (g.nube && g.nube.h) || {};
  const grading = g.grading || {};
  const secs = sectionsOf(g).map((s, i) => ({
    id: s.nube_id || null, local: s.id, nombre: s.name, plantilla: s.type,
    porcentaje: weightOf(grading, s), minimo: s.type === 'sino' ? secMin(s) : null, orden: i,
  }));
  const conf = { materia: g.subject || '', retardo: Number(grading.retardo ?? 100), escala: Number(grading.escala ?? 10), secs: secs.map(({ id, ...x }) => x) };
  const hash = huella(JSON.stringify(conf));
  if (h.__secciones === hash && secs.every((s) => s.id)) return null;
  return { ...conf, secs, hash };
}

// Días por subir. Requiere secciones y alumnos ya con id de la nube.
// dias: [{ key, hash, seccion, fecha, valores }] · trabajos: [{ key, hash, sec, fecha, works }]
// borrar: llaves de días que ya no existen (se mandan vacíos) · olvidar: llaves de secciones quitadas.
export function planDias(g) {
  const h = (g.nube && g.nube.h) || {};
  const secs = sectionsOf(g);
  const out = { dias: [], trabajos: [], olvidar: [] };
  const vistos = new Set();
  for (const s of secs) {
    if (!s.nube_id) continue;
    for (const fecha of Object.keys(g[s.id] || {}).sort()) {
      const key = `${s.id}|${fecha}`;
      vistos.add(key);
      if (s.type === 'trab') {
        const works = worksOf(g, fecha, s.id);
        const hash = huella(JSON.stringify(works.map((w) => [w.title || '', valores(g, 'trab', w.vals)])));
        if (h[key] !== hash || works.some((w) => !w.nube_id)) out.trabajos.push({ key, hash, sec: s, fecha, works });
      } else {
        const v = valores(g, s.type, g[s.id][fecha]);
        const hash = huella(JSON.stringify(v));
        if (h[key] !== hash) out.dias.push({ key, hash, seccion: s.nube_id, fecha, valores: v });
      }
    }
  }
  // Días que se subieron antes y ya no existen en el celular (se borró el día o sus trabajos).
  const porId = Object.fromEntries(secs.map((s) => [s.id, s]));
  for (const key of Object.keys(h)) {
    if (key === '__secciones' || vistos.has(key)) continue;
    const [sid, fecha] = key.split('|');
    const s = porId[sid];
    if (!s || !s.nube_id) { out.olvidar.push(key); continue; }
    if (s.type === 'trab') out.trabajos.push({ key, hash: null, sec: s, fecha, works: [] });
    else out.dias.push({ key, hash: null, seccion: s.nube_id, fecha, valores: {} });
  }
  return out;
}

// Días de un trabajo ya con id: uno por trabajo.
export const diasDeTrabajos = (g, t) => t.works.map((w) => ({ seccion: t.sec.nube_id, fecha: t.fecha, trabajo: w.nube_id, valores: valores(g, 'trab', w.vals) }));

// ---------- Bajar de la nube (dispositivo nuevo) ----------
const normNombre = (t) => String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
const valorLocal = (type, v) => (type === 'asist' || type === 'sino' ? v : Number(v));

// Llena un grupo leído del Excel (g) con lo que hay en la nube para su clase. Lo que traía el Excel en
// columnas de la app se reemplaza por la nube. Alumnos inscritos que no están en el Excel (agregados en
// la app) se agregan al final. Al terminar, todo queda marcado como ya subido.
export function armarGrupo(g, { clase, secciones, inscripciones, trabajos, capturas, cierres = [] }) {
  g.subject = clase.materia || undefined;
  if (!g.subject) delete g.subject;
  const ordenadas = [...secciones].sort((a, b) => a.orden - b.orden);
  // Datos que traía el Excel en columnas de la app: se quitan (la nube manda).
  for (const s of g.sections || []) delete g[s.id];
  g.sections = ordenadas.map((s) => ({
    id: s.plantilla === 'asist' ? 'records' : `n_${s.id.replace(/-/g, '').slice(0, 12)}`,
    name: s.nombre, type: s.plantilla, nube_id: s.id, ...(s.plantilla === 'sino' ? { min: s.minimo || 4 } : {}),
  }));
  const local = Object.fromEntries(g.sections.map((s) => [s.nube_id, s]));
  g.grading = { w: {}, retardo: clase.retardo, escala: clase.escala };
  for (const s of ordenadas) if (Number(s.porcentaje) > 0) g.grading.w[local[s.id].id] = Number(s.porcentaje);
  g.sections.forEach((s) => (g[s.id] = {}));
  // Alumnos: por número de lista y nombre; si no, por nombre.
  const filaDe = {};
  const libres = [...g.students];
  for (const i of inscripciones) {
    const nombre = normNombre(i.nombre);
    let s = libres.find((x) => String(x.no) === String(i.numero_lista) && normNombre(x.name) === nombre)
      || libres.find((x) => normNombre(x.name) === nombre);
    if (!s) {
      const row = Math.max(g.headerRow, ...g.students.map((x) => x.row)) + 1;
      s = { row, no: i.numero_lista, name: i.nombre, added: true };
      g.students.push(s);
    } else libres.splice(libres.indexOf(s), 1);
    s.alumno_id = i.alumno_id;
    filaDe[i.alumno_id] = s.row;
  }
  // Trabajos por sección y fecha, en orden.
  const works = {};
  for (const t of [...trabajos].sort((a, b) => a.orden - b.orden)) {
    const s = local[t.seccion_id];
    if (!s) continue;
    const lista = (g[s.id][t.fecha] = g[s.id][t.fecha] || []);
    const w = { title: t.titulo || '', vals: {}, nube_id: t.id };
    lista.push(w);
    works[t.id] = w;
  }
  for (const c of capturas) {
    const s = local[c.seccion_id];
    const row = filaDe[c.alumno_id];
    if (!s || row == null) continue;
    const v = valorLocal(s.type, c.valor);
    if (s.type === 'trab') { if (works[c.trabajo_id]) works[c.trabajo_id].vals[row] = v; }
    else (g[s.id][c.fecha] = g[s.id][c.fecha] || {})[row] = v;
  }
  // Calificación semestral (cómo se saca y exámenes finales) y cierres, con alumnos → filas.
  const aFila = (obj) => Object.fromEntries(Object.entries(obj || {}).filter(([a]) => filaDe[a] != null).map(([a, v]) => [filaDe[a], v]));
  const cfg = clase.config_semestral && Object.keys(clase.config_semestral).length ? clase.config_semestral : null;
  if (cfg) g.semestral = { ...cfg, final: aFila(cfg.final) };
  g.cierres = [...cierres].sort((a, b) => (a.cerrado_en < b.cerrado_en ? -1 : 1)).map((c) => ({
    tipo: c.tipo, parcial: c.parcial ?? null, cerrado_en: c.cerrado_en,
    calificaciones: Object.fromEntries(Object.entries(aFila(c.calificaciones)).map(([row, v]) => [row, v && v.partes
      ? { ...v, partes: Object.fromEntries(Object.entries(v.partes).map(([k, x]) => [(local[k] || {}).id || k, x])) } : v])),
    reabierto_en: c.reabierto_en || null, nube_id: c.id, nube_reabierto: !!c.reabierto_en,
  }));
  // Todo lo que llegó ya está en la nube: se marca como subido.
  g.nube = { clase_id: clase.id, grupo_id: clase.grupo_id, h: {} };
  if (cfg) g.nube.h.__semestral = huella(JSON.stringify(cfg));
  const ps = planSecciones(g);
  if (ps) g.nube.h.__secciones = ps.hash;
  const p = planDias(g);
  [...p.dias, ...p.trabajos].forEach((x) => { if (x.hash) g.nube.h[x.key] = x.hash; });
  return g;
}
