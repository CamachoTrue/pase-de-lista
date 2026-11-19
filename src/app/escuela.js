// ---------- Mi escuela (pantalla del director) ----------
// El director no necesita grupos ni listas: al entrar ve su escuela con tres pestañas.
//   Resumen   → el mismo panel de /director (alumnos, asistencia, promedios, riesgo, maestros).
//   Maestros  → invitar, ver invitaciones y dar o quitar acceso a cada maestro.
//   Parciales → fechas de los parciales del semestre para toda la escuela.
// En la demostración de director (/?demo=director) la escuela es de ejemplo y no se guarda nada.
import '../director/director.css';
import { crearPanel, escuelaDeEjemplo, exportarEscuela, leerEscuela } from '../director/panel.js';
import { errorLegible, nube } from '../lib/nube.js';
import { DEMO } from '../lib/demo.js';
import { $, escapeHtml, toast } from './util.js';
import { openCuenta } from './menu.js';

const ROL = { director: 'Director', maestro: 'Maestro' };
const fecha = (iso) => new Date(iso).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
const enlace = (token) => `${location.origin}/?invitacion=${token}`;
const ejemplo = () => DEMO === 'director';
const main = () => $('#esc-main');

let pestaña = 'resumen';
let panel = null;
let perfil = null; // { escuela_id, escuela, clave, id }
let cargado = { resumen: false, maestros: false, parciales: false };

export async function openEscuela() {
  $('#screen-escuela').hidden = false;
  if (!main().dataset.listo) armar();
  cargado = { resumen: false, maestros: false, parciales: false };
  perfil = null;
  await mostrarPestaña(pestaña);
}

function armar() {
  main().dataset.listo = '1';
  main().innerHTML = `
    <div data-pane="resumen">
      <div class="esc-acciones"><button class="btn" data-esc="actualizar">↻ Actualizar</button>
        <button class="btn primary" data-esc="exportar" hidden>Exportar Excel</button></div>
      <div data-inicio></div>
      <div class="dir-vista-app" data-panel><p class="hint">Cargando…</p></div>
    </div>
    <div data-pane="maestros" hidden></div>
    <div data-pane="parciales" hidden></div>`;
  panel = crearPanel(main().querySelector('[data-panel]'));
}

async function datosPerfil() {
  if (ejemplo()) return { escuela_id: 'ejemplo', escuela: 'Preparatoria de Ejemplo', clave: 'EJEMPLO', id: 'yo' };
  if (perfil) return perfil;
  const sb = await nube();
  const { data, error } = await sb.rpc('mi_perfil');
  if (error) throw error;
  perfil = data && data[0];
  if (!perfil || perfil.rol !== 'director') throw new Error('Esta pantalla es para directores.');
  return perfil;
}

async function mostrarPestaña(p) {
  pestaña = p;
  document.querySelectorAll('#esc-tabs [data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === p));
  main().querySelectorAll('[data-pane]').forEach((d) => { d.hidden = d.dataset.pane !== p; });
  if (cargado[p]) return;
  const pane = main().querySelector(`[data-pane="${p}"]`);
  try {
    const yo = await datosPerfil();
    $('#esc-nombre').textContent = yo.escuela || 'Mi escuela';
    $('#esc-sub').textContent = yo.clave ? `Clave ${yo.clave}` : '';
    if (p === 'resumen') await cargarResumen(yo);
    if (p === 'maestros') await cargarMaestros(pane, yo);
    if (p === 'parciales') await cargarParciales(pane, yo);
    cargado[p] = true;
  } catch (e) {
    console.error(e);
    const txt = /fetch|network/i.test(String(e && e.message)) || !navigator.onLine
      ? 'Sin internet. El panel de tu escuela necesita conexión; vuelve a intentarlo cuando tengas.' : errorLegible(e);
    (p === 'resumen' ? main().querySelector('[data-panel]') : pane).innerHTML = `<div class="cuenta-card"><b>No se pudo cargar</b><small>${escapeHtml(txt)}</small>
      <button class="btn" data-esc="actualizar">Reintentar</button></div>`;
  }
}

// ---------- Resumen ----------
async function cargarResumen(yo) {
  const box = main().querySelector('[data-panel]');
  box.innerHTML = '<p class="hint">Cargando datos de la escuela…</p>';
  const sb = ejemplo() ? null : await nube();
  const datos = ejemplo() ? await escuelaDeEjemplo() : await leerEscuela(sb, yo.escuela_id);
  const inicio = main().querySelector('[data-inicio]');
  main().querySelector('[data-esc="exportar"]').hidden = !datos.clases.length;
  if (datos.clases.length) {
    inicio.innerHTML = '';
    box.hidden = false;
    panel.mostrar(datos);
    return;
  }
  // Escuela sin clases todavía: en vez de un panel en ceros, qué sigue.
  const [{ count: parciales }, { count: invitados }] = await Promise.all([
    sb.from('semestres').select('id', { count: 'exact', head: true }).eq('escuela_id', yo.escuela_id),
    sb.from('invitaciones').select('id', { count: 'exact', head: true }).eq('escuela_id', yo.escuela_id),
  ]);
  inicio.innerHTML = primerosPasos({ maestros: datos.maestros.length || invitados || 0, parciales: parciales || 0 });
  box.hidden = true;
}

// Escuela recién dada de alta: qué sigue, paso por paso.
function primerosPasos({ maestros, parciales }) {
  const paso = (hecho, txt, accion = '') => `<li class="${hecho ? 'hecho' : ''}"><span>${hecho ? '✓' : ''}</span><div>${txt}${accion}</div></li>`;
  return `<section class="cuenta-card esc-pasos">
      <b>Primeros pasos</b>
      <small>Cuando tus maestros empiecen a pasar lista, aquí aparecerá el resumen de la escuela.</small>
      <ol>
        ${paso(maestros > 0, 'Invita a tus maestros. Les llega un enlace para crear su cuenta.', ' <button class="link-btn" data-ir="maestros">Ir a Maestros ›</button>')}
        ${paso(parciales > 0, 'Pon las fechas de los 3 parciales del semestre (valen para todos).', ' <button class="link-btn" data-ir="parciales">Ir a Parciales ›</button>')}
        ${paso(false, 'Cada maestro entra a la app, sube su Excel y pasa lista como siempre.')}
        ${paso(false, 'Aquí verás alumnos, asistencia, promedios y quién está en riesgo, sin pedirle nada a nadie.')}
      </ol>
    </section>`;
}

// ---------- Maestros: invitar y dar o quitar acceso ----------
const EJ_USUARIOS = [
  { id: 'yo', nombre: 'Mtra. Laura Gómez', correo: 'direccion@prepaejemplo.mx', rol: 'director', activo: true },
  { id: 'm1', nombre: 'Prof. Jorge Méndez', correo: 'jorge.mendez@prepaejemplo.mx', rol: 'maestro', activo: true },
  { id: 'm2', nombre: 'Mtra. Sofía Navarro', correo: 'sofia.navarro@prepaejemplo.mx', rol: 'maestro', activo: true },
  { id: 'm3', nombre: 'Prof. Ricardo Soto', correo: 'ricardo.soto@prepaejemplo.mx', rol: 'maestro', activo: false },
];

async function cargarMaestros(pane, yo) {
  let usuarios, invitaciones, cobro;
  if (ejemplo()) {
    usuarios = EJ_USUARIOS;
    invitaciones = [{ id: 'i1', correo: 'rosa.leyva@prepaejemplo.mx', token: 'ejemplo', vence_en: new Date(Date.now() + 5 * 864e5).toISOString() }];
    cobro = { cuentas_contratadas: 8, cuentas_usadas: 3 };
  } else {
    const sb = await nube();
    const [c, inv, us] = await Promise.all([
      sb.rpc('datos_cobro'),
      sb.from('invitaciones').select('id, correo, token, vence_en').eq('escuela_id', yo.escuela_id).is('usada_en', null).order('creada_en', { ascending: false }),
      sb.from('usuarios').select('id, nombre, correo, rol, activo').eq('escuela_id', yo.escuela_id).order('nombre'),
    ]);
    if (us.error) throw us.error;
    cobro = c.data && c.data[0];
    invitaciones = (inv.data || []).filter((i) => new Date(i.vence_en) > new Date());
    usuarios = us.data || [];
  }
  const ocupadas = cobro ? cobro.cuentas_usadas + invitaciones.length : null;
  const libres = cobro ? cobro.cuentas_contratadas - ocupadas : null;
  const yoId = ejemplo() ? 'yo' : (await (await nube()).auth.getSession()).data.session?.user.id;
  pane.innerHTML = `
    <section class="cuenta-sec">
      <h3>Invitar a un maestro</h3>
      <p class="hint">${cobro ? `Usas <b>${ocupadas} de ${cobro.cuentas_contratadas}</b> cuentas (las invitaciones pendientes también cuentan).` : ''}
        Escribe su correo y mándale el enlace por WhatsApp; con él crea su contraseña.</p>
      <form class="cuenta-form invitar" data-form="invitar">
        <label class="field">Correo del maestro<input name="correo" type="email" autocomplete="off" required placeholder="maestro@correo.com"></label>
        <button class="btn primary" type="submit" ${libres != null && libres <= 0 ? 'disabled' : ''}>Crear invitación</button>
        ${libres != null && libres <= 0 ? '<p class="hint">Ya usas todas tus cuentas. Quita el acceso a alguien que ya no esté o pide más cuentas.</p>' : ''}
        <p class="form-error" data-error></p>
        <div data-enlace></div>
      </form>
    </section>
    ${invitaciones.length ? `<section class="cuenta-sec"><h3>Invitaciones pendientes</h3><ul class="pick-list cuenta-lista">
      ${invitaciones.map((i) => `<li data-token="${i.token}"><span class="nm">${escapeHtml(i.correo)}<small>vence ${fecha(i.vence_en)} · aún no crea su cuenta</small></span>
        <button class="link-btn" data-esc="enviar">Enviar</button>
        <button class="link-btn danger-text" data-esc="quitar-inv" data-id="${i.id}">Quitar</button></li>`).join('')}
      </ul></section>` : ''}
    <section class="cuenta-sec">
      <h3>Cuentas de tu escuela</h3>
      <p class="hint">Si un maestro ya no está, <b>quítale el acceso</b>: deja de entrar a la cuenta y se libera su lugar. Sus calificaciones se quedan en la escuela.</p>
      <ul class="pick-list cuenta-lista esc-usuarios">
      ${usuarios.map((u) => `<li class="${u.activo ? '' : 'inactivo'}"><span class="nm">${escapeHtml(u.nombre)}${u.id === yoId ? ' (tú)' : ''}
          <small>${escapeHtml(u.correo)} · ${ROL[u.rol] || u.rol}${u.activo ? '' : ' · <b class="danger-text">sin acceso</b>'}</small></span>
          ${u.id === yoId || u.rol !== 'maestro' ? '' : u.activo
            ? `<button class="link-btn danger-text" data-esc="acceso" data-id="${u.id}" data-activo="0">Quitar acceso</button>`
            : `<button class="link-btn" data-esc="acceso" data-id="${u.id}" data-activo="1" ${libres != null && libres <= 0 ? 'disabled title="No tienes cuentas libres"' : ''}>Devolver acceso</button>`}
        </li>`).join('')}
      </ul>
    </section>`;
}

// ---------- Parciales ----------
async function cargarParciales(pane, yo) {
  const filas = [1, 2, 3].map((n) => `<div class="per-row"><b>Parcial ${n}</b><input type="date" name="p${n}i" aria-label="Inicio del parcial ${n}"><span>a</span><input type="date" name="p${n}f" aria-label="Fin del parcial ${n}"></div>`).join('');
  pane.innerHTML = `<section class="cuenta-sec">
      <h3>Parciales del semestre</h3>
      <p class="hint">Valen para <b>todos los maestros</b> de tu escuela: cada asistencia y calificación cae sola en su parcial, y el maestro ya no tiene que ponerlas.</p>
      <form class="cuenta-form" data-form="periodos">
        <label class="per-sem">Semestre <input name="semestre" required placeholder="Ago–Dic 2026" maxlength="40"></label>
        ${filas}
        <button class="btn primary" type="submit">Guardar parciales</button>
        <p class="form-error" data-error></p>
      </form>
    </section>`;
  const f = pane.querySelector('form');
  if (ejemplo()) {
    const y = new Date().getFullYear();
    f.elements.semestre.value = `Ago–Dic ${y}`;
    [['p1', '08-25', '09-26'], ['p2', '09-29', '10-31'], ['p3', '11-03', '12-12']].forEach(([k, i, fin]) => { f.elements[`${k}i`].value = `${y}-${i}`; f.elements[`${k}f`].value = `${y}-${fin}`; });
    return;
  }
  const sb = await nube();
  const { data } = await sb.from('semestres').select('id, nombre, parciales(numero, inicio, fin)').eq('escuela_id', yo.escuela_id).order('inicio', { ascending: false }).limit(1);
  const s = data && data[0];
  if (!s) return;
  f.elements.semestre.value = s.nombre;
  for (const p of s.parciales) {
    if (!f.elements[`p${p.numero}i`]) continue;
    f.elements[`p${p.numero}i`].value = p.inicio;
    f.elements[`p${p.numero}f`].value = p.fin;
  }
}

async function guardarPeriodos(sb, f, escuela) {
  const pars = [1, 2, 3].map((n) => ({ numero: n, inicio: f.elements[`p${n}i`].value, fin: f.elements[`p${n}f`].value })).filter((p) => p.inicio && p.fin);
  if (!pars.length) throw new Error('Escribe las fechas de al menos un parcial.');
  if (pars.some((p) => p.fin < p.inicio)) throw new Error('Cada parcial debe terminar después de empezar.');
  const nombre = f.elements.semestre.value.trim();
  const inicio = pars.map((p) => p.inicio).sort()[0];
  const fin = pars.map((p) => p.fin).sort().at(-1);
  const { data: sem, error } = await sb.from('semestres').upsert({ escuela_id: escuela, nombre, inicio, fin }, { onConflict: 'escuela_id,nombre' }).select('id').single();
  if (error) throw error;
  const r = await sb.from('parciales').upsert(pars.map((p) => ({ ...p, escuela_id: escuela, semestre_id: sem.id })), { onConflict: 'semestre_id,numero' });
  if (r.error) throw r.error;
  const quitar = [1, 2, 3, 4, 5, 6].filter((n) => !pars.some((p) => p.numero === n));
  await sb.from('parciales').delete().eq('semestre_id', sem.id).in('numero', quitar);
}

// ---------- Enlaces de invitación ----------
async function copiar(texto) {
  try { await navigator.clipboard.writeText(texto); toast('Enlace copiado'); } catch { toast('Copia el enlace manualmente'); }
}
async function compartir(token) {
  const url = enlace(token);
  const text = 'Te invito a la app para pasar lista y calificar de nuestra escuela. Crea tu cuenta aquí:';
  if (navigator.share) {
    try { await navigator.share({ title: 'Invitación', text, url }); return; } catch { /* cancelado: se copia */ }
  }
  await copiar(url);
}

// ---------- Eventos ----------
$('#esc-tabs').addEventListener('click', (e) => {
  const b = e.target.closest('[data-tab]');
  if (b) mostrarPestaña(b.dataset.tab);
});
$('#esc-cuenta').addEventListener('click', openCuenta);

main().addEventListener('click', async (e) => {
  const ir = e.target.closest('[data-ir]');
  if (ir) { mostrarPestaña(ir.dataset.ir); return; }
  const b = e.target.closest('[data-esc]');
  if (!b) return;
  const act = b.dataset.esc;
  if (act === 'actualizar') { cargado[pestaña] = false; await mostrarPestaña(pestaña); return; }
  if (act === 'exportar') {
    b.disabled = true;
    try { await exportarEscuela(panel.datos(), $('#esc-nombre').textContent); } catch (ex) { console.error(ex); toast('No se pudo exportar. Revisa tu conexión.'); }
    b.disabled = false;
    return;
  }
  if (ejemplo()) { toast('Demostración: aquí no se guarda nada'); return; }
  const token = (b.closest('[data-token]') || {}).dataset?.token;
  if (act === 'copiar') copiar(enlace(token));
  if (act === 'enviar') compartir(token);
  if (act === 'quitar-inv' || act === 'acceso') {
    const { confirmBulk } = await import('./dialogs.js');
    const sb = await nube();
    let error;
    if (act === 'quitar-inv') {
      if (!(await confirmBulk('¿Quitar esta invitación?', 'El enlace dejará de funcionar y se libera la cuenta.', 'Sí, quitar', true))) return;
      ({ error } = await sb.from('invitaciones').delete().eq('id', b.dataset.id));
    } else {
      const activo = b.dataset.activo === '1';
      if (!activo && !(await confirmBulk('¿Quitar el acceso a este maestro?', 'Ya no podrá entrar a su cuenta ni guardar en la nube. Sus calificaciones se quedan en la escuela y puedes devolverle el acceso después.', 'Sí, quitar acceso', true))) return;
      ({ error } = await sb.from('usuarios').update({ activo }).eq('id', b.dataset.id));
      if (!error) toast(activo ? 'Acceso devuelto' : 'Acceso quitado');
    }
    if (error) { toast(errorLegible(error)); return; }
    cargado.maestros = false;
    await mostrarPestaña('maestros');
  }
});

main().addEventListener('submit', async (e) => {
  e.preventDefault();
  const f = e.target;
  const err = f.querySelector('[data-error]');
  const btn = f.querySelector('[type=submit]');
  if (ejemplo()) { toast('Demostración: aquí no se guarda nada'); return; }
  err.textContent = '';
  btn.disabled = true;
  try {
    const sb = await nube();
    const yo = await datosPerfil();
    if (f.dataset.form === 'invitar') {
      const correo = f.elements.correo.value.trim().toLowerCase();
      const { data, error } = await sb.from('invitaciones').insert({ escuela_id: yo.escuela_id, correo, rol: 'maestro' }).select('token').single();
      if (error) throw error;
      cargado.maestros = false;
      await mostrarPestaña('maestros');
      const box = main().querySelector('[data-pane="maestros"] [data-enlace]');
      box.innerHTML = `<div class="cuenta-enlace" data-token="${data.token}">
          <b>Invitación lista para ${escapeHtml(correo)}</b>
          <small>Mándale este enlace. Sirve una sola vez, solo con ese correo, y vence en 7 días.</small>
          <input readonly value="${enlace(data.token)}">
          <div class="row"><button type="button" class="btn" data-esc="copiar">Copiar</button><button type="button" class="btn primary" data-esc="enviar">Enviar por WhatsApp</button></div>
        </div>`;
    } else if (f.dataset.form === 'periodos') {
      await guardarPeriodos(sb, f, yo.escuela_id);
      toast('Parciales guardados para toda la escuela');
    }
  } catch (ex) {
    console.error(ex);
    const m = String(ex.message || '');
    err.textContent = /cuentas contratadas/.test(m) ? 'Ya usas todas tus cuentas. Quita el acceso a alguien que ya no esté o pide más cuentas.' : errorLegible(ex);
  } finally {
    btn.disabled = false;
  }
});
