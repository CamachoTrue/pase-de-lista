// ---------- Bajar los datos de la cuenta a un dispositivo nuevo ----------
// Descarga el Excel original, las clases del maestro con sus secciones, alumnos y capturas, y los
// cuadernillos; arma los grupos igual que en el otro celular y deja todo marcado como ya subido.
import { DEFAULT_CODES, normalizeGroup, parseWorkbook } from '../lib/excel-core.js';
import { armarGrupo, huella } from '../lib/sync-plan.js';
import { nube } from '../lib/nube.js';
import { setState, save, state, todayISO, fixGrading } from './state.js';
import { bufToB64, loadExcel } from './util.js';
import { bookDB } from './book-db.js';
import { ASIST_ID } from '../lib/excel-core.js';

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

// Clases guardadas en la cuenta (para ofrecer bajarlas).
export async function clasesEnLaNube() {
  const sb = await nube();
  const { data: { session } } = await sb.auth.getSession();
  if (!session) return 0;
  const { count } = await sb.from('clases').select('id', { count: 'exact', head: true }).eq('maestro_id', session.user.id);
  return count || 0;
}

export async function bajarDeLaNube(avance = () => {}) {
  const sb = await nube();
  const { data: { session } } = await sb.auth.getSession();
  if (!session) throw new Error('Inicia sesión primero');
  const uid = session.user.id;
  const { data: perfil, error: ep } = await sb.rpc('mi_perfil');
  if (ep) throw ep;
  const carpeta = `${perfil[0].escuela_id}/${uid}`;
  const st = sb.storage.from('archivos');

  avance('Bajando tu Excel…');
  const { data: lista } = await st.list(`${carpeta}/excel`);
  if (!lista || !lista.length) throw new Error('Tu cuenta aún no tiene tu Excel. Ábrelo primero en el celular donde capturas, con internet, para que se guarde.');
  const archivo = lista[0].name;
  const { data: blob, error: eb } = await st.download(`${carpeta}/excel/${archivo}`);
  if (eb) throw eb;
  const buf = await blob.arrayBuffer();

  avance('Bajando tus clases…');
  const clases = await todas(() => sb.from('clases').select('id, grupo_id, materia, retardo, escala, config_semestral, cambios, grupos(nombre)').eq('maestro_id', uid).order('creado_en'));
  const ids = clases.map((c) => c.id);
  const secciones = ids.length ? await todas(() => sb.from('secciones').select('id, clase_id, nombre, plantilla, porcentaje, minimo, orden').in('clase_id', ids)) : [];
  const secIds = secciones.map((s) => s.id);
  const inscripciones = ids.length ? await todas(() => sb.from('inscripciones').select('grupo_id, alumno_id, numero_lista, baja, alumnos(nombre)').in('grupo_id', clases.map((c) => c.grupo_id))) : [];
  const trabajos = secIds.length ? await todas(() => sb.from('trabajos').select('id, seccion_id, fecha, titulo, orden').in('seccion_id', secIds)) : [];
  avance('Bajando calificaciones y asistencias…');
  const capturas = ids.length ? await todas(() => sb.from('capturas').select('clase_id, seccion_id, alumno_id, fecha, trabajo_id, valor').in('clase_id', ids).order('id')) : [];

  const cierres = ids.length ? await todas(() => sb.from('cierres').select('id, clase_id, tipo, parcial, calificaciones, cerrado_en, reabierto_en').in('clase_id', ids)) : [];
  avance('Armando tus grupos…');
  const ExcelJS = await loadExcel();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const claseDe = Object.fromEntries(clases.map((c) => [c.grupos.nombre, c]));
  const groups = parseWorkbook(wb, DEFAULT_CODES);
  for (const g of groups) {
    normalizeGroup(g);
    g.grading = fixGrading(null);
    const c = claseDe[g.sheet];
    if (!c) continue; // hoja sin clase en la nube: queda como en el Excel
    armarGrupo(g, {
      clase: c,
      secciones: secciones.filter((s) => s.clase_id === c.id),
      inscripciones: inscripciones.filter((i) => i.grupo_id === c.grupo_id).map((i) => ({ alumno_id: i.alumno_id, numero_lista: i.numero_lista, nombre: i.alumnos.nombre, baja: i.baja })),
      trabajos: trabajos.filter((t) => secciones.some((s) => s.id === t.seccion_id && s.clase_id === c.id)),
      capturas: capturas.filter((x) => x.clase_id === c.id),
      cierres: cierres.filter((x) => x.clase_id === c.id),
    });
  }
  const fileB64 = bufToB64(buf);
  setState({
    fileName: archivo, fileB64, groups, codes: { ...DEFAULT_CODES, ...((state && state.codes) || {}) },
    ui: { group: 0, date: todayISO(), tab: ASIST_ID }, nube_usuario: uid,
    nube_excel: huella(`${archivo}|${fileB64.length}|${fileB64.slice(-200)}`),
  });
  save(true);

  // Cuadernillos (si fallan, lo demás ya quedó).
  avance('Bajando cuadernillos…');
  const { data: cuadernillos } = await sb.from('cuadernillos').select('clase_id, nombre, ruta, tamano, pagina').in('clase_id', ids.length ? ids : ['00000000-0000-0000-0000-000000000000']);
  for (const k of cuadernillos || []) {
    const g = groups.find((x) => x.nube && x.nube.clase_id === k.clase_id);
    if (!g) continue;
    try {
      const { data: pdf } = await st.download(k.ruta);
      const rec = { name: k.nombre, size: k.tamano, data: await pdf.arrayBuffer(), page: k.pagina || 1, added: Date.now() };
      await bookDB.put(g.sheet, rec);
      g.nube.pdf = huella(`${rec.name}|${rec.size}|${rec.added}`);
      g.nube.pdf_nube = `${k.nombre}|${k.tamano}`;
    } catch (e) { console.error(e); }
  }
  save(true);
  return groups.filter((g) => g.nube).length;
}
