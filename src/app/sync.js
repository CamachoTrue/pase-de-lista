// ---------- Sincronización con la cuenta ----------
// Todo se guarda primero en el celular (así funciona sin internet). Si hay sesión, lo que cambió se
// sube solo unos segundos después, al volver el internet y al abrir la app.
import { alGuardar, fixGrading, save, state } from './state.js';
import { dispositivo, errorLegible, haySesion, nube } from '../lib/nube.js';
import { armarGrupo, diasDeTrabajos, faltaImportar, hayQueBajar, huella, planDias, planSecciones, resumenClase } from '../lib/sync-plan.js';
import { ASIST_ID, DEFAULT_CODES, normalizeGroup } from '../lib/excel-core.js';
import { b64ToBuf, toast } from './util.js';
import { bookDB } from './book-db.js';

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
    // Suscripción vencida o escuela desactivada: no se sube nada (los datos siguen en el celular).
    const p0 = perfil[0];
    if (!p0.activa || (p0.vence_el && p0.vence_el < new Date().toISOString().slice(0, 10))) { estado.fase = 'vencida'; return; }
    state.nube_usuario = session.user.id;
    await bajarPeriodos(sb, perfil[0].escuela_id);
    await sincronizarAjustes(sb, session.user.id);
    for (const g of state.groups) await subirGrupo(sb, g);
    for (const g of state.groups) await subirSemestral(sb, g, perfil[0].escuela_id, session.user.id);
    const carpeta = `${perfil[0].escuela_id}/${session.user.id}`;
    await subirExcel(sb, carpeta);
    await subirCuadernillos(sb, carpeta, perfil[0].escuela_id);
    // Primero se sube lo de este dispositivo y después se baja lo que cambió en otro: así lo capturado
    // aquí nunca se pierde y, si los dos cambiaron el mismo día, gana el último que se guardó en la nube.
    const n = await bajarCambios(sb, session.user.id);
    for (const g of state.groups) await subirResumen(sb, g, perfil[0].escuela_id);
    if (n) toast(n === 1 ? 'Se actualizó 1 grupo con lo capturado en otro dispositivo' : `Se actualizaron ${n} grupos con lo capturado en otro dispositivo`);
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

// ---------- Parciales, calificación semestral y cierres ----------
// Fechas de parciales que definió la escuela (el director): mandan sobre las del maestro.
async function bajarPeriodos(sb, escuela) {
  const hoy = new Date().toISOString().slice(0, 10);
  const { data: sems, error } = await sb.from('semestres').select('id, nombre, inicio, fin, parciales(numero, inicio, fin)').eq('escuela_id', escuela).order('inicio', { ascending: false });
  if (error) throw error;
  if (!sems || !sems.length) {
    if (state.periodos && state.periodos.origen === 'escuela') { state.periodos = null; save(true); }
    return;
  }
  const s = sems.find((x) => x.inicio <= hoy && hoy <= x.fin) || sems[0];
  const periodos = { semestre: s.nombre, id: s.id, origen: 'escuela',
    parciales: [...s.parciales].sort((a, b) => a.numero - b.numero).map((p) => ({ numero: p.numero, inicio: p.inicio, fin: p.fin })) };
  if (JSON.stringify(periodos) !== JSON.stringify(state.periodos)) { state.periodos = periodos; save(true); }
}

// ---------- Ajustes de la cuenta: valores de cada estado y parciales propios del maestro ----------
// Lo que este dispositivo manda a la cuenta. Si los parciales los puso la escuela, se conservan los
// propios que ya estaban en la cuenta.
const ajustesLocales = (remoto) => ({
  codes: state.codes,
  periodos: state.periodos && state.periodos.origen === 'escuela' ? (remoto.periodos ?? null) : (state.periodos || null),
});
async function sincronizarAjustes(sb, uid) {
  const { data, error } = await sb.from('usuarios').select('ajustes, ajustes_en').eq('id', uid).single();
  if (error) throw error;
  const remoto = data.ajustes || {};
  const hayRemoto = !!data.ajustes_en;
  const h = huella(JSON.stringify(ajustesLocales(remoto)));
  // Primera vez en este dispositivo: si la cuenta ya tiene ajustes, mandan los de la cuenta.
  const cambioAqui = state.nube_ajustes ? h !== state.nube_ajustes : !hayRemoto;
  if (cambioAqui) {
    const { data: r, error: e } = await sb.from('usuarios').update({ ajustes: ajustesLocales(remoto) }).eq('id', uid).select('ajustes_en').single();
    if (e) throw e;
    state.nube_ajustes = h;
    state.nube_ajustes_en = r.ajustes_en;
    save(true);
  } else if (hayRemoto && data.ajustes_en !== state.nube_ajustes_en) {
    // Otro dispositivo los cambió: se usan aquí.
    if (remoto.codes) state.codes = { ...DEFAULT_CODES, ...remoto.codes };
    if (!(state.periodos && state.periodos.origen === 'escuela')) state.periodos = remoto.periodos || null;
    state.nube_ajustes = huella(JSON.stringify(ajustesLocales(remoto)));
    state.nube_ajustes_en = data.ajustes_en;
    save(true);
    window.dispatchEvent(new Event('nube-cambio'));
  }
}

// { fila: x } → { alumno_id: x }
const porAlumno = (g, obj) => {
  const id = Object.fromEntries(g.students.filter((s) => s.alumno_id).map((s) => [s.row, s.alumno_id]));
  return Object.fromEntries(Object.entries(obj || {}).filter(([row]) => id[row]).map(([row, v]) => [id[row], v]));
};

// En la nube, el desglose por sección de un cierre usa el id de la sección en la nube (igual en todos los celulares).
const conPartesNube = (g, cal) => {
  const nubeDe = Object.fromEntries((g.sections || []).map((s) => [s.id, s.nube_id || s.id]));
  return Object.fromEntries(Object.entries(cal || {}).map(([row, v]) => [row, v && v.partes
    ? { ...v, partes: Object.fromEntries(Object.entries(v.partes).map(([k, x]) => [nubeDe[k] || k, x])) } : v]));
};

async function subirSemestral(sb, g, escuela, uid) {
  const clase = g.nube && g.nube.clase_id;
  if (!clase) return;
  // Cómo se saca la semestral (y los exámenes finales), en la clase.
  if (g.semestral) {
    const cfg = { ...g.semestral, final: porAlumno(g, g.semestral.final) };
    const hs = huella(JSON.stringify(cfg));
    if (g.nube.h.__semestral !== hs) {
      const { error } = await sb.from('clases').update({ config_semestral: cfg }).eq('id', clase);
      if (error) throw error;
      g.nube.h.__semestral = hs;
      save(true);
    }
  }
  // Cierres nuevos y reaperturas.
  for (const c of g.cierres || []) {
    if (!c.nube_id) {
      const { data, error } = await sb.from('cierres').insert({
        escuela_id: escuela, clase_id: clase, tipo: c.tipo, parcial: c.parcial || null, semestre: (state.periodos && state.periodos.semestre) || '',
        calificaciones: porAlumno(g, conPartesNube(g, c.calificaciones)), cerrado_por: uid, cerrado_en: c.cerrado_en,
      }).select('id').single();
      if (error) throw error;
      c.nube_id = data.id;
      save(true);
    }
    if (c.reabierto_en && !c.nube_reabierto) {
      const { error } = await sb.from('cierres').update({ reabierto_en: c.reabierto_en, reabierto_por: uid }).eq('id', c.nube_id);
      if (error) throw error;
      c.nube_reabierto = true;
      save(true);
    }
  }
}

// Resumen de la clase para el panel del director (solo si cambió).
async function subirResumen(sb, g, escuela) {
  const clase = g.nube && g.nube.clase_id;
  if (!clase) return;
  const { datos, ultima } = resumenClase(g, state.periodos);
  const hs = huella(JSON.stringify([datos, ultima]));
  if (g.nube.h.__resumen === hs) return;
  const { error } = await sb.from('resumenes').upsert({ clase_id: clase, escuela_id: escuela, datos, ultima_captura: ultima, actualizado_en: new Date().toISOString() }, { onConflict: 'clase_id' });
  if (error) throw error;
  g.nube.h.__resumen = hs;
  save(true);
}

// ---------- Archivos: Excel original y cuadernillos (Supabase Storage, carpeta propia) ----------
export const nombreArchivo = (n) => String(n || 'lista.xlsx').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9._-]+/g, '_').slice(-80);
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function subirExcel(sb, carpeta) {
  if (!state.fileB64) return;
  const marca = huella(`${state.fileName}|${state.fileB64.length}|${state.fileB64.slice(-200)}`);
  if (state.nube_excel === marca) return;
  const st = sb.storage.from('archivos');
  const nombre = nombreArchivo(state.fileName);
  const { error } = await st.upload(`${carpeta}/excel/${nombre}`, new Blob([b64ToBuf(state.fileB64)], { type: XLSX }), { upsert: true, contentType: XLSX });
  if (error) throw error;
  // Solo queda el Excel actual.
  const { data: viejos } = await st.list(`${carpeta}/excel`);
  const sobran = (viejos || []).filter((f) => f.name !== nombre).map((f) => `${carpeta}/excel/${f.name}`);
  if (sobran.length) await st.remove(sobran);
  state.nube_excel = marca;
  save(true);
}

async function subirCuadernillos(sb, carpeta, escuela) {
  const st = sb.storage.from('archivos');
  for (const g of state.groups) {
    const clase = g.nube && g.nube.clase_id;
    if (!clase) continue;
    let rec = null;
    try { rec = await bookDB.get(g.sheet); } catch { /* sin IndexedDB */ }
    const ruta = `${carpeta}/cuadernillos/${clase}.pdf`;
    if (rec) {
      const marca = huella(`${rec.name}|${rec.size}|${rec.added || ''}`);
      if (g.nube.pdf === marca) continue;
      const up = await st.upload(ruta, new Blob([rec.data], { type: 'application/pdf' }), { upsert: true, contentType: 'application/pdf' });
      if (up.error) throw up.error;
      const r = await sb.from('cuadernillos').upsert({ escuela_id: escuela, clase_id: clase, nombre: rec.name, ruta, tamano: rec.size, pagina: rec.page || 1 }, { onConflict: 'clase_id' });
      if (r.error) throw r.error;
      g.nube.pdf = marca;
      g.nube.pdf_nube = `${rec.name}|${rec.size}`;
      save(true);
    } else if (g.nube.pdf) {
      await st.remove([ruta]);
      await sb.from('cuadernillos').delete().eq('clase_id', clase);
      delete g.nube.pdf;
      delete g.nube.pdf_nube;
      save(true);
    }
  }
}

// ---------- Bajar lo que cambió en otro dispositivo ----------
const PAGINA = 1000;
async function todas(consulta) {
  const out = [];
  for (let desde = 0; ; desde += PAGINA) {
    const { data, error } = await consulta().range(desde, desde + PAGINA - 1);
    if (error) throw error;
    out.push(...data);
    if (data.length < PAGINA) return out;
  }
}

// Ediciones hechas en este dispositivo (para no reemplazar un grupo mientras alguien captura).
let ediciones = 0;
const ocupado = () => !!document.querySelector('dialog[open]:not(#cuenta)');

async function cambiosEnLaNube(sb, uid) {
  const { data, error } = await sb.from('clases').select('id, cambios').eq('maestro_id', uid);
  if (error) throw error;
  return Object.fromEntries(data.map((c) => [c.id, c.cambios]));
}

async function bajarCambios(sb, uid) {
  if (!state.groups.some((g) => g.nube && g.nube.clase_id)) return 0;
  const cambios = await cambiosEnLaNube(sb, uid);
  const este = dispositivo();
  let n = 0;
  for (const g of state.groups) {
    if (!hayQueBajar(g, cambios[g.nube && g.nube.clase_id], este)) continue;
    if (await bajarClase(sb, g)) n++;
  }
  if (n) window.dispatchEvent(new Event('nube-cambio'));
  return n;
}

// Baja una clase completa y la pone en el grupo (el mismo objeto). Regresa false si mientras bajaba
// alguien capturó algo aquí: se deja para la siguiente vuelta.
async function bajarClase(sb, g) {
  const id = g.nube.clase_id;
  const antes = ediciones;
  const { data: clase, error } = await sb.from('clases').select('id, grupo_id, materia, retardo, escala, config_semestral, cambios').eq('id', id).maybeSingle();
  if (error) throw error;
  if (!clase) return false;
  const secciones = await todas(() => sb.from('secciones').select('id, clase_id, nombre, plantilla, porcentaje, minimo, orden').eq('clase_id', id).order('id'));
  const secIds = secciones.map((s) => s.id);
  const inscripciones = await todas(() => sb.from('inscripciones').select('grupo_id, alumno_id, numero_lista, alumnos(nombre)').eq('grupo_id', clase.grupo_id).is('baja', null).order('alumno_id'));
  const trabajos = secIds.length ? await todas(() => sb.from('trabajos').select('id, seccion_id, fecha, titulo, orden').in('seccion_id', secIds).order('id')) : [];
  const capturas = await todas(() => sb.from('capturas').select('clase_id, seccion_id, alumno_id, fecha, trabajo_id, valor').eq('clase_id', id).order('id'));
  const cierres = await todas(() => sb.from('cierres').select('id, clase_id, tipo, parcial, calificaciones, cerrado_en, reabierto_en').eq('clase_id', id).order('id'));
  const { data: cuad } = await sb.from('cuadernillos').select('nombre, ruta, tamano, pagina').eq('clase_id', id).maybeSingle();
  if (ediciones !== antes || ocupado()) return false;

  armarGrupo(g, {
    clase, secciones, trabajos, capturas, cierres,
    inscripciones: inscripciones.map((i) => ({ alumno_id: i.alumno_id, numero_lista: i.numero_lista, nombre: i.alumnos.nombre })),
  });
  g.grading = fixGrading(g.grading);
  normalizeGroup(g, g.grading);
  if (state.groups[state.ui.group] === g && !(g.sections || []).some((s) => s.id === state.ui.tab)) state.ui.tab = ASIST_ID;
  save(true);

  // Cuadernillo: si otro dispositivo lo puso, lo cambió o lo quitó.
  try {
    if (cuad) {
      const marca = `${cuad.nombre}|${cuad.tamano}`;
      if (g.nube.pdf_nube !== marca) {
        const { data: pdf, error: e } = await sb.storage.from('archivos').download(cuad.ruta);
        if (e) throw e;
        const rec = { name: cuad.nombre, size: cuad.tamano, data: await pdf.arrayBuffer(), page: cuad.pagina || 1, added: Date.now() };
        await bookDB.put(g.sheet, rec);
        g.nube.pdf = huella(`${rec.name}|${rec.size}|${rec.added}`);
        g.nube.pdf_nube = marca;
      }
    } else if (g.nube.pdf) {
      await bookDB.del(g.sheet);
      delete g.nube.pdf;
      delete g.nube.pdf_nube;
    }
  } catch (e) { console.error(e); }
  save(true);
  return true;
}

// Revisión ligera (una consulta): al volver a la app y cada 30 s mientras está abierta. Si otro
// dispositivo cambió algo, se sincroniza completo (sube lo de aquí y baja lo de allá).
async function revisar() {
  if (!state || !state.groups || !haySesion() || !navigator.onLine || corriendo || document.hidden) return;
  if (!state.nube_usuario) return; // este dispositivo aún no se ha ligado a la cuenta
  try {
    const sb = await nube();
    const { data: { session } } = await sb.auth.getSession();
    if (!session || state.nube_usuario !== session.user.id) return;
    const [cambios, { data: yo }] = await Promise.all([
      cambiosEnLaNube(sb, session.user.id),
      sb.from('usuarios').select('ajustes_en').eq('id', session.user.id).maybeSingle(),
    ]);
    const este = dispositivo();
    const ajustes = yo && yo.ajustes_en && yo.ajustes_en !== state.nube_ajustes_en;
    if (ajustes || state.groups.some((g) => hayQueBajar(g, cambios[g.nube && g.nube.clase_id], este))) sincronizar();
  } catch { /* sin conexión: se intenta en la siguiente vuelta */ }
}
setInterval(revisar, 30000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) setTimeout(revisar, 500); });

// Cada cambio guardado programa una subida; también al volver el internet y al abrir la app.
alGuardar(() => { ediciones++; programar(); });
window.addEventListener('online', () => programar(1000));
window.addEventListener('cuadernillo-cambio', () => programar(1500));
if (haySesion()) setTimeout(() => programar(500), 2500);
