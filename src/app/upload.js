import { ASIST_ID, DEFAULT_CODES, normalizeGroup, parseWorkbook } from '../lib/excel-core.js';
import { render } from './main-view.js';
import { fixGrading, save, setState, state, todayISO } from './state.js';
import { $, bufToB64, loadExcel, toast } from './util.js';

// ---------- Subir archivo ----------
$('#file-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  try {
    const buf = await file.arrayBuffer();
    const wb = new (await loadExcel()).Workbook();
    await wb.xlsx.load(buf);
    const codes = (state && state.codes) || { ...DEFAULT_CODES };
    // Si el grupo ya existía (misma hoja), conserva sus secciones, porcentajes y materia.
    const old = (state && state.groups) || [];
    const prevOf = (sheet) => old.find((o) => o.sheet === sheet);
    const groups = parseWorkbook(wb, codes, (sheet) => (prevOf(sheet) || {}).sections);
    if (!groups.length) {
      toast('No encontré una columna “NOMBRE” en ninguna hoja');
      return;
    }
    groups.forEach((g) => {
      const prev = prevOf(g.sheet);
      normalizeGroup(g);
      g.grading = fixGrading(prev && prev.grading);
      if (prev && prev.subject) g.subject = prev.subject; // la materia también se conserva
      // La clase en la nube también: se vuelve a subir todo (la base de datos no duplica nada).
      if (prev && prev.nube && prev.nube.clase_id) g.nube = { clase_id: prev.nube.clase_id, grupo_id: prev.nube.grupo_id, h: {} };
    });
    setState({ fileName: file.name, fileB64: bufToB64(buf), groups, codes: { ...DEFAULT_CODES, ...codes }, ui: { group: 0, date: todayISO(), tab: ASIST_ID }, nube_usuario: state && state.nube_usuario,
      // Parciales y marcas de la cuenta no dependen del Excel: se conservan.
      periodos: state && state.periodos, nube_ajustes: state && state.nube_ajustes, nube_ajustes_en: state && state.nube_ajustes_en });
    save();
    const total = groups.reduce((n, g) => n + g.students.length, 0);
    toast(`${groups.length} grupo(s), ${total} alumnos cargados`);
    render();
  } catch (err) {
    console.error(err);
    toast('No se pudo leer el archivo. ¿Es un .xlsx?');
  }
});
