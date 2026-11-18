// ---------- Sincronización con la cuenta ----------
// Todo se guarda primero en el celular (así funciona sin internet). Si hay sesión, lo que cambió se
// sube solo unos segundos después, al volver el internet y al abrir la app.
import { alGuardar, save, state } from './state.js';
import { errorLegible, haySesion, nube } from '../lib/nube.js';
import { diasDeTrabajos, faltaImportar, planDias, planSecciones } from '../lib/sync-plan.js';

const LOTE = 60; // días por llamada
export const estado = { fase: 'inactivo', ultimo: null, error: '', subidos: 0 };
const avisos = [];
export const alCambiarEstado = (fn) => avisos.push(fn);
const avisar = () => avisos.forEach((fn) => fn(estado));

let timer = null;
let corriendo = false;
let otraVez = false;
export function programar(ms = 4000) {
  if (!haySesion()) return;
  clearTimeout(timer);
  timer = setTimeout(sincronizar, ms);
}

export async function sincronizar() {
  if (!state || !state.groups || !haySesion()) return;
  if (!navigator.onLine) { estado.fase = 'sin-conexion'; avisar(); return; }
  if (corriendo) { otraVez = true; return; }
  corriendo = true;
  estado.fase = 'subiendo';
  estado.subidos = 0;
  avisar();
  try {
    const sb = await nube();
    const { data: { session } } = await sb.auth.getSession();
    if (!session) { estado.fase = 'inactivo'; return; }
    // Los datos de este celular son de una sola cuenta: si entra otra, no se mezclan.
    if (state.nube_usuario && state.nube_usuario !== session.user.id) { estado.fase = 'otra-cuenta'; return; }
    const { data: perfil, error } = await sb.rpc('mi_perfil');
    if (error) throw error;
    if (!perfil || !perfil[0] || !perfil[0].escuela_id) { estado.fase = 'sin-escuela'; return; }
    state.nube_usuario = session.user.id;
    for (const g of state.groups) await subirGrupo(sb, g);
    estado.fase = 'listo';
    estado.ultimo = Date.now();
    estado.error = '';
  } catch (e) {
    console.error(e);
    estado.fase = /fetch|network/i.test(String(e && e.message)) ? 'sin-conexion' : 'error';
    estado.error = errorLegible(e);
  } finally {
    corriendo = false;
    save(true);
    avisar();
    if (otraVez) { otraVez = false; programar(800); }
  }
}

const rpc = async (sb, fn, args) => {
  const { data, error } = await sb.rpc(fn, args);
  if (error) throw error;
  return data;
};

async function subirGrupo(sb, g) {
  g.nube = g.nube || {};
  g.nube.h = g.nube.h || {};
  // 1. Clase y alumnos (reconoce a los que ya están en la escuela).
  if (faltaImportar(g)) {
    const r = await rpc(sb, 'importar_grupo', {
      p_grupo: g.sheet, p_materia: g.subject || '', p_alumnos: g.students.map((s) => ({ numero: s.no, nombre: s.name })),
    });
    g.nube.clase_id = r.clase_id;
    g.nube.grupo_id = r.grupo_id;
    g.students.forEach((s, i) => { s.alumno_id = r.alumnos[i]; });
    save(true);
  }
  // 2. Secciones y porcentajes.
  const ps = planSecciones(g);
  if (ps) {
    const ids = await rpc(sb, 'guardar_secciones', {
      p_clase: g.nube.clase_id, p_materia: ps.materia, p_retardo: ps.retardo, p_escala: ps.escala,
      p_secciones: ps.secs.map(({ id, local, nombre, plantilla, porcentaje, minimo, orden }) => ({ id, local, nombre, plantilla, porcentaje, minimo, orden })),
    });
    g.sections.forEach((s) => { s.nube_id = ids[s.id]; });
    g.nube.h.__secciones = ps.hash;
    save(true);
  }
  // 3. Días.
  const p = planDias(g);
  p.olvidar.forEach((k) => delete g.nube.h[k]);
  const listos = [];
  if (p.trabajos.length) {
    const res = await rpc(sb, 'guardar_trabajos', {
      p_dias: p.trabajos.map((t) => ({ seccion: t.sec.nube_id, fecha: t.fecha, trabajos: t.works.map((w) => ({ id: w.nube_id || null, titulo: w.title || '' })) })),
    });
    p.trabajos.forEach((t, i) => { t.works.forEach((w, j) => { w.nube_id = res[i].ids[j]; }); });
    save(true);
    p.trabajos.forEach((t) => listos.push({ key: t.key, hash: t.hash, dias: diasDeTrabajos(g, t) }));
  }
  p.dias.forEach((d) => listos.push({ key: d.key, hash: d.hash, dias: [{ seccion: d.seccion, fecha: d.fecha, valores: d.valores }] }));
  // En lotes: si se corta el internet a la mitad, lo ya subido queda marcado y se sigue después.
  for (let i = 0; i < listos.length; i += LOTE) {
    const lote = listos.slice(i, i + LOTE);
    const dias = lote.flatMap((x) => x.dias);
    if (dias.length) await rpc(sb, 'guardar_dias', { p_dias: dias });
    lote.forEach((x) => { if (x.hash == null) delete g.nube.h[x.key]; else g.nube.h[x.key] = x.hash; });
    estado.subidos += lote.length;
    save(true);
    avisar();
  }
}

// Cada cambio guardado programa una subida; también al volver el internet y al abrir la app.
alGuardar(() => programar());
window.addEventListener('online', () => programar(1000));
if (haySesion()) setTimeout(() => programar(500), 2500);
