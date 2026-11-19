// ---------- Panel de la escuela (lo usan /director y la pantalla «Mi escuela» de la app) ----------
// Lee los resúmenes que suben los celulares de los maestros (tabla "resumenes") y arma la vista de la
// escuela: indicadores, clases, maestros, alumnos en riesgo y la ficha de cada alumno con todas sus materias.
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const r1 = (n) => Math.round(n * 10) / 10;
const avg = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : null);
const hoy = () => new Date().toISOString().slice(0, 10);
const dias = (iso) => (iso ? Math.floor((Date.parse(hoy()) - Date.parse(iso)) / 864e5) : null);
const hace = (iso) => { const n = dias(iso); return n == null || n < 0 ? '' : n === 0 ? ' <small>(hoy)</small>' : ` <small>(hace ${n} d)</small>`; };
const fecha = (iso) => (iso ? new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' }) : '—');
const ESTADO = { rojo: 'En riesgo', amarillo: 'Observación', verde: 'Bien', none: 'Sin datos' };
const punto = (e) => `<span class="dir-dot ${e || 'none'}" title="${ESTADO[e || 'none']}"></span>`;

async function todas(consulta) {
  const out = [];
  for (let d = 0; ; d += 1000) {
    const { data, error } = await consulta().range(d, d + 999);
    if (error) throw error;
    out.push(...data);
    if (data.length < 1000) return out;
  }
}

// Filas de la escuela desde la base de datos (lo que se puede ver lo decide RLS).
export async function leerEscuela(sb, escuela) {
  const [clases, resumenes, alumnos, maestros] = await Promise.all([
    todas(() => sb.from('clases').select('id, materia, maestro_id, grupo_id, grupos(nombre)').eq('escuela_id', escuela)),
    todas(() => sb.from('resumenes').select('clase_id, datos, ultima_captura, actualizado_en').eq('escuela_id', escuela)),
    todas(() => sb.from('alumnos').select('id, nombre, clave').eq('escuela_id', escuela)),
    todas(() => sb.from('usuarios').select('id, nombre, correo, rol, activo').eq('escuela_id', escuela)),
  ]);
  return armar(escuela, { clases, resumenes, alumnos, maestros });
}

// Escuela de ejemplo (datos inventados) con la misma forma.
export async function escuelaDeEjemplo() {
  const { escuelaEjemplo } = await import('./ejemplo.js');
  return armar('ejemplo', escuelaEjemplo());
}

// Filas de la base de datos (o de la escuela de ejemplo) → lo que pinta el panel.
export function armar(escuela, { clases, resumenes, alumnos, maestros }) {
  const res = Object.fromEntries(resumenes.map((r) => [r.clase_id, r]));
  const maestro = Object.fromEntries(maestros.map((m) => [m.id, m]));
  return {
    escuela,
    alumnos: Object.fromEntries(alumnos.map((a) => [a.id, a])),
    maestros: maestros.filter((m) => m.rol === 'maestro' || clases.some((c) => c.maestro_id === m.id)),
    clases: clases.map((c) => {
      const r = res[c.id];
      const al = r ? r.datos.alumnos || {} : {};
      const escala = (r && r.datos.escala) || 10;
      const filas = Object.values(al);
      const cuenta = (e) => filas.filter((x) => x.e === e).length;
      return {
        id: c.id, grupo: c.grupos ? c.grupos.nombre : '', materia: c.materia, maestro: (maestro[c.maestro_id] || {}).nombre || '—',
        maestro_id: c.maestro_id, escala, alumnos: al, n: filas.length,
        promedio: avg(filas.map((x) => x.f).filter((v) => v != null)), asistencia: avg(filas.map((x) => x.a).filter((v) => v != null)),
        rojo: cuenta('rojo'), amarillo: cuenta('amarillo'), verde: cuenta('verde'),
        ultima: r ? r.ultima_captura : null, actualizado: r ? r.actualizado_en : null, semestre: r ? r.datos.semestre : '',
      };
    }).sort((a, b) => (a.grupo + a.materia).localeCompare(b.grupo + b.materia)),
  };
}

// Calificación en escala de 10 (las clases pueden calificar sobre 10 o sobre 100).
const sobre10 = (v, escala) => (v == null ? null : (v * 10) / (escala || 10));
const calif = (v) => (v == null ? '—' : r1(v));

// Monta el panel en un contenedor. mostrar(datos) lo pinta; los clics y la búsqueda se quedan dentro.
export function crearPanel(root) {
  let datos = null;
  let vista = 'resumen';
  let detalle = null; // clase_id o alumno_id abierto
  function pintar() {
    const d = datos;
    const todos = d.clases.flatMap((c) => Object.entries(c.alumnos).map(([id, x]) => ({ id, ...x, c })));
    const unicos = new Set(todos.map((x) => x.id));
    const riesgo = new Set(todos.filter((x) => x.e === 'rojo').map((x) => x.id));
    const prom = avg(todos.map((x) => sobre10(x.f, x.c.escala)).filter((v) => v != null));
    const asis = avg(todos.map((x) => x.a).filter((v) => v != null));
    const atrasadas = d.clases.filter((c) => c.ultima == null || dias(c.ultima) > 7).length;
    const tabs = [['resumen', 'Clases'], ['maestros', 'Por maestro'], ['riesgo', `En riesgo (${riesgo.size})`], ['alumno', 'Buscar alumno']];
    // Cada indicador es un botón que lleva a su detalle.
    const tile = (v, txt, ir, cls = '') => `<button class="gr-tile dir-tile ${cls}" data-vista="${ir}"><b>${v}</b><span>${txt}</span><i>Ver ›</i></button>`;
    let html = `<section class="dir-tiles">
        ${tile(unicos.size, 'Alumnos', 'alumno')}
        ${tile(d.maestros.filter((m) => m.activo).length, 'Maestros', 'maestros')}
        ${tile(d.clases.length, 'Clases', 'resumen')}
        ${tile(asis == null ? '—' : `${r1(asis)} %`, 'Asistencia promedio', 'resumen')}
        ${tile(prom == null ? '—' : r1(prom), 'Promedio general (sobre 10)', 'resumen')}
        ${tile(riesgo.size, 'Alumnos en riesgo', 'riesgo', riesgo.size ? 'rojo' : '')}
        ${tile(atrasadas, 'Clases sin captura en 7 días', 'maestros', atrasadas ? 'amar' : '')}
      </section>
      <nav class="dir-tabs">${tabs.map(([k, t]) => `<button class="per-tab ${vista === k ? 'on' : ''}" data-vista="${k}">${t}</button>`).join('')}</nav>
      <section class="dir-vista">`;
    if (vista === 'resumen') html += detalle && d.clases.find((c) => c.id === detalle) ? vistaClase(d.clases.find((c) => c.id === detalle)) : vistaClases(d);
    if (vista === 'maestros') html += vistaMaestros(d);
    if (vista === 'riesgo') html += vistaRiesgo(d, todos);
    if (vista === 'alumno') html += vistaAlumno(d, todos);
    root.innerHTML = html + '</section>';
    const q = root.querySelector('#dir-buscar');
    if (q) { q.focus(); q.setSelectionRange(q.value.length, q.value.length); }
  }

  function vistaClases(d) {
    if (!d.clases.length) return '<p class="hint">Aún no hay clases. Aparecen cuando los maestros entran con su cuenta y sincronizan.</p>';
    return `<p class="hint">Toca una clase para ver a sus alumnos, y el nombre de un alumno para ver todas sus materias. Los datos los suben los celulares de los maestros al sincronizar.</p>
      <div class="preview-scroll"><table class="sheet dir-tabla"><thead><tr><th>Grupo</th><th>Materia</th><th>Maestro</th><th>Alumnos</th>
      <th>Promedio</th><th>Asistencia</th><th>Semáforo</th><th>Última captura</th><th></th></tr></thead><tbody>
      ${d.clases.map((c) => `<tr class="dir-click" data-clase="${c.id}"><td><b>${esc(c.grupo)}</b></td><td>${esc(c.materia || '—')}</td><td>${esc(c.maestro)}</td>
        <td class="v">${c.n}</td><td class="v">${c.promedio == null ? '—' : `${calif(c.promedio)}<small> /${c.escala}</small>`}</td>
        <td class="v">${c.asistencia == null ? '—' : `${r1(c.asistencia)} %`}</td>
        <td class="dir-sem"><span class="rojo">${c.rojo}</span><span class="amarillo">${c.amarillo}</span><span class="verde">${c.verde}</span></td>
        <td class="${c.ultima == null || dias(c.ultima) > 7 ? 'dir-atraso' : ''}">${c.ultima ? `${fecha(c.ultima)}${hace(c.ultima)}` : 'sin datos'}</td><td class="dir-chev">Ver ›</td></tr>`).join('')}
      </tbody></table></div>`;
  }

  function vistaClase(c) {
    const filas = Object.entries(c.alumnos).map(([id, x]) => ({ id, nombre: (datos.alumnos[id] || {}).nombre || '—', ...x }))
      .sort((a, b) => ({ rojo: 0, amarillo: 1, verde: 2, none: 3 }[a.e] - { rojo: 0, amarillo: 1, verde: 2, none: 3 }[b.e]) || a.nombre.localeCompare(b.nombre));
    const np = Math.max(0, ...filas.map((x) => (x.p || []).length));
    return `<button class="link-btn" data-volver>‹ Todas las clases</button>
      <h2 class="dir-h2">${esc(c.grupo)} · ${esc(c.materia || 'sin materia')} <small>· ${esc(c.maestro)} · sobre ${c.escala}${c.semestre ? ` · ${esc(c.semestre)}` : ''}</small></h2>
      <div class="preview-scroll"><table class="sheet dir-tabla"><thead><tr><th></th><th>Alumno</th><th>Calif.</th><th>Asistencia</th><th>Faltas</th>
      ${Array.from({ length: np }, (_, i) => `<th>P${i + 1}</th>`).join('')}${np ? '<th>Semestral</th>' : ''}<th>Motivo</th></tr></thead><tbody>
      ${filas.map((x) => `<tr><td>${punto(x.e)}</td><td class="dir-click" data-alumno="${x.id}">${esc(x.nombre)}</td><td class="v">${calif(x.f)}</td>
        <td class="v">${x.a == null ? '—' : `${r1(x.a)} %`}</td><td class="v">${x.F ?? 0}</td>
        ${(x.p || []).map((v) => `<td class="v">${calif(v)}</td>`).join('')}${np ? `<td class="v"><b>${calif(x.s)}</b></td>` : ''}
        <td class="dir-motivo">${esc((x.r || []).join(' · '))}</td></tr>`).join('')}
      </tbody></table></div>`;
  }

  function vistaMaestros(d) {
    const filas = d.maestros.map((m) => {
      const cs = d.clases.filter((c) => c.maestro_id === m.id);
      const ultima = cs.map((c) => c.ultima).filter(Boolean).sort().at(-1) || null;
      const alumnos = cs.reduce((n, c) => n + c.n, 0);
      const rojo = cs.reduce((n, c) => n + c.rojo, 0);
      return { m, cs, ultima, alumnos, rojo };
    }).sort((a, b) => (dias(b.ultima) ?? 999) - (dias(a.ultima) ?? 999));
    return `<p class="hint">Primero quien lleva más días sin capturar.</p>
      <div class="preview-scroll"><table class="sheet dir-tabla"><thead><tr><th>Maestro</th><th>Clases</th><th>Alumnos</th><th>En riesgo</th><th>Última captura</th></tr></thead><tbody>
      ${filas.map((x) => `<tr><td><b>${esc(x.m.nombre)}</b><small>${esc(x.m.correo)}${x.m.activo ? '' : ' · desactivado'}</small></td>
        <td>${x.cs.map((c) => `${esc(c.grupo)}${c.materia ? ` ${esc(c.materia)}` : ''}`).join(', ') || '—'}</td><td class="v">${x.alumnos}</td>
        <td class="v ${x.rojo ? 'dir-rojo' : ''}">${x.rojo}</td>
        <td class="${x.ultima == null || dias(x.ultima) > 7 ? 'dir-atraso' : ''}">${x.ultima ? `${fecha(x.ultima)}${hace(x.ultima)}` : 'sin capturas'}</td></tr>`).join('')}
      </tbody></table></div>`;
  }

  function vistaRiesgo(d, todos) {
    const rojos = todos.filter((x) => x.e === 'rojo').sort((a, b) => ((d.alumnos[a.id] || {}).nombre || '').localeCompare((d.alumnos[b.id] || {}).nombre || ''));
    if (!rojos.length) return '<p class="hint">Ningún alumno en riesgo. 🎉</p>';
    return `<p class="hint">Alumnos con calificación reprobatoria o asistencia menor a 80 % en alguna materia.</p>
      <div class="preview-scroll"><table class="sheet dir-tabla"><thead><tr><th>Alumno</th><th>Grupo · materia</th><th>Calif.</th><th>Asistencia</th><th>Motivo</th></tr></thead><tbody>
      ${rojos.map((x) => `<tr><td class="dir-click" data-alumno="${x.id}"><b>${esc((d.alumnos[x.id] || {}).nombre)}</b></td>
        <td>${esc(x.c.grupo)} · ${esc(x.c.materia || '—')}<small>${esc(x.c.maestro)}</small></td><td class="v">${calif(x.f)}<small> /${x.c.escala}</small></td>
        <td class="v">${x.a == null ? '—' : `${r1(x.a)} %`}</td><td class="dir-motivo">${esc((x.r || []).join(' · '))}</td></tr>`).join('')}
      </tbody></table></div>`;
  }

  let busqueda = '';
  const norm = (t) => String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
  function vistaAlumno(d, todos) {
    let html = `<input id="dir-buscar" class="dialog-search" placeholder="Nombre o clave del alumno…" value="${esc(busqueda)}">`;
    if (detalle && d.alumnos[detalle]) {
      const a = d.alumnos[detalle];
      const filas = todos.filter((x) => x.id === detalle);
      html += `<h2 class="dir-h2">${esc(a.nombre)} <small>· ${esc(a.clave)}</small></h2>
        <div class="preview-scroll"><table class="sheet dir-tabla"><thead><tr><th></th><th>Grupo · materia</th><th>Maestro</th><th>Calif.</th><th>Asistencia</th><th>Faltas</th><th>Semestral</th><th>Motivo</th></tr></thead><tbody>
        ${filas.map((x) => `<tr><td>${punto(x.e)}</td><td>${esc(x.c.grupo)} · ${esc(x.c.materia || '—')}</td><td>${esc(x.c.maestro)}</td>
          <td class="v">${calif(x.f)}<small> /${x.c.escala}</small></td><td class="v">${x.a == null ? '—' : `${r1(x.a)} %`}</td><td class="v">${x.F ?? 0}</td>
          <td class="v">${x.s == null ? '—' : calif(x.s)}</td><td class="dir-motivo">${esc((x.r || []).join(' · '))}</td></tr>`).join('') || '<tr><td colspan="8">Sin materias con datos.</td></tr>'}
        </tbody></table></div>`;
      return html;
    }
    const q = norm(busqueda).trim();
    if (q.length >= 2) {
      const hits = Object.values(d.alumnos).filter((a) => norm(a.nombre).includes(q) || norm(a.clave).includes(q)).slice(0, 30);
      html += `<ul class="pick-list">${hits.map((a) => `<li class="dir-click" data-alumno="${a.id}"><span class="nm">${esc(a.nombre)}</span><span class="no">${esc(a.clave)}</span></li>`).join('') || '<li class="empty">Nadie coincide.</li>'}</ul>`;
    } else html += '<p class="hint">Escribe al menos 2 letras.</p>';
    return html;
  }


  root.addEventListener('click', (e) => {
    const t = e.target.closest('[data-vista],[data-clase],[data-alumno],[data-volver]');
    if (!t) return;
    if (t.dataset.vista) { vista = t.dataset.vista; detalle = null; }
    else if (t.dataset.clase) { vista = 'resumen'; detalle = t.dataset.clase; }
    else if (t.dataset.alumno) { vista = 'alumno'; detalle = t.dataset.alumno; }
    else if (t.hasAttribute('data-volver')) detalle = null;
    pintar();
    window.scrollTo(0, 0);
  });
  root.addEventListener('input', (e) => {
    if (e.target.id !== 'dir-buscar') return;
    busqueda = e.target.value;
    detalle = null;
    pintar();
  });

  return { mostrar: (d) => { datos = d; pintar(); }, datos: () => datos };
}

// Excel de la escuela (para el director): se arma solo al pedirlo.
export async function exportarEscuela(datos, nombre) {
  if (!datos) return;
  const { excelEscuela, entregar } = await import('./excel-escuela.js');
  const ExcelJS = await cargarExcelJS();
  const buf = await excelEscuela(ExcelJS, datos, nombre);
  const limpio = String(nombre || 'escuela').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]+/g, '_').slice(0, 40);
  await entregar(buf, `Resumen_${limpio}_${hoy()}.xlsx`);
}
let excelListo = null;
const cargarExcelJS = () => (excelListo = excelListo || new Promise((res, rej) => {
  if (window.ExcelJS) { res(window.ExcelJS); return; }
  const sc = document.createElement('script');
  sc.src = '/vendor/exceljs.min.js';
  sc.onload = () => res(window.ExcelJS);
  sc.onerror = () => { excelListo = null; rej(new Error('No se pudo cargar Excel')); };
  document.head.appendChild(sc);
}));
