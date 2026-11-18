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
