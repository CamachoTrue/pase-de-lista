import { ASIST_ID, DEFAULT_GRADING, sectionsOf } from '../lib/excel-core.js';
import { toast } from './util.js';

export const LABELS = { A: 'Asistencia', F: 'Falta', R: 'Retardo', J: 'Justificado' };
// Opciones de captura según la plantilla de la sección: valor guardado, texto del botón y color (clase A/F/R/J/N).
export const OPTIONS = {
  sino: [
    { v: 'S', btn: 'Sí', color: 'A' },
    { v: 'N', btn: 'No', color: 'F' },
  ],
  trab: [
    { v: 0, btn: '0', color: 'F' },
    { v: 50, btn: '50', color: 'R' },
    { v: 100, btn: '100', color: 'A' },
  ],
  // Disciplina: todos tienen 10 sin capturar nada; solo se guarda el 0.
  disc: [
    { v: 10, btn: '10', color: 'N' },
    { v: 0, btn: '0', color: 'F' },
  ],
};
export const DISC_DEFAULT = 10;
// Valores del botón "restantes" según la plantilla.
export const FILL_REST = { sino: ['S', 'N'], trab: [0] };
// Secciones del grupo actual (cada maestro las agrega, quita y renombra).
export const secs = () => sectionsOf(group());
export const secById = (id) => secs().find((x) => x.id === id);
export const tab = () => (state.ui.tab && secById(state.ui.tab) ? state.ui.tab : ASIST_ID);
export const typeOf = (id) => (secById(id) || { type: 'asist' }).type;
export const optionOf = (id, v) => {
  const sec = secById(id) || { type: 'asist', name: 'Asistencia' };
  if (sec.type === 'asist') return v ? { v, label: LABELS[v], color: v } : null;
  if (sec.type === 'num') return v != null ? { v, label: `${sec.name}: ${v}`, color: 'N' } : null;
  const o = OPTIONS[sec.type].find((x) => x.v === v);
  return o ? { ...o, label: `${sec.name}: ${o.btn}` } : null;
};
// Datos de una sección para el día seleccionado.
// Trabajos: puede haber varios por día; se usa el trabajo seleccionado (state.ui.work).
export const works = (create) => {
  const g = group();
  const id = tab();
  g[id] = g[id] || {};
  if (!g[id][state.ui.date] && create) g[id][state.ui.date] = [];
  return g[id][state.ui.date] || [];
};
export const workIdx = () => Math.max(0, Math.min(state.ui.work || 0, Math.max(works(false).length - 1, 0)));
export const curWork = (create) => {
  const ws = works(create);
  const i = workIdx();
  if (!ws[i] && create) ws[i] = { title: '', vals: {} };
  return ws[i] || null;
};
export const dayData = (id, create) => {
  if (typeOf(id) === 'trab') { const w = curWork(create); return w ? w.vals : {}; }
  const g = group();
  g[id] = g[id] || {};
  if (!g[id][state.ui.date] && create) g[id][state.ui.date] = {};
  return g[id][state.ui.date] || {};
};
// Porcentajes de un grupo, completos (con los valores por defecto que falten).
export const fixGrading = (gr) => JSON.parse(JSON.stringify({ ...DEFAULT_GRADING, ...(gr || {}), w: { ...DEFAULT_GRADING.w, ...((gr || {}).w || {}) } }));
// Nombre del grupo para mostrar: hoja del Excel y, si el maestro la puso, la materia (p. ej. "V205 · Ética").
export const gname = (g) => (g.subject ? `${g.sheet} · ${g.subject}` : g.sheet);
// En el selector: con materia se omite el número de alumnos para que quepa el nombre.
export const glabel = (g) => (g.subject ? gname(g) : `${g.sheet} · ${g.students.length} alumnos`);
export const STORE_KEY = 'paseDeLista.v1';

// ---------- Estado ----------
// { fileName, fileB64, groups, codes, ui: { group, date } }
export let state = load();
// Reemplaza todo el estado (al subir otro archivo).
export function setState(s) { state = s; }
// Estado de la pantalla que comparten varios módulos.
export const ui = { view: 'card', cardIdx: 0, busy: false, studentRow: null };

export function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* almacenamiento no disponible */ }
  return null;
}
export function save() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
  catch (e) { toast('No se pudo guardar en el dispositivo'); }
}

export const todayISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};
export const group = () => state.groups[state.ui.group];
export const dayRecords = (create) => {
  const g = group();
  if (!g.records[state.ui.date] && create) g.records[state.ui.date] = {};
  return g.records[state.ui.date] || {};
};
export const fmtDate = (iso) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
};
