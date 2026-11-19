// ---------- Mi cuenta (Supabase) ----------
// Entrar con correo y contraseña, mantener la sesión, invitar maestros (director) y dar de alta
// escuelas (superadmin). Por ahora las listas y calificaciones siguen guardándose en el celular;
// la sincronización con la cuenta llega en la etapa 4.
import { $, escapeHtml, toast } from './util.js';
import { confirmBulk } from './dialogs.js';
import { errorLegible, guardarAcceso, guardarPaquete, guardarRol, nube, recordar, rolGuardado, setRecordar } from '../lib/nube.js';
import { alCambiarEstado, estado, programar, sincronizar } from './sync.js';
import { save, setVistaApp, state } from './state.js';
import { DEMO, PERFIL_DEMO } from '../lib/demo.js';
import { bajarDeLaNube, clasesEnLaNube } from './bajar.js';
import { render, vistaEscuela } from './main-view.js';

const ROL = { superadmin: 'Administrador general', director: 'Director', maestro: 'Maestro' };
const body = () => $('#cuenta-body');
let perfil = null; // el de la sesión, para "Eliminar mi cuenta"
const enlace = (token) => `${location.origin}/?invitacion=${token}`;

// Directo al registro (desde el aviso «Crear mi cuenta gratis»).
export function openRegistro() {
  $('#cuenta').showModal();
  $('#cuenta-sub').textContent = '';
  renderRegistro();
}

export async function openCuenta() {
  $('#cuenta').showModal();
  if (DEMO) { renderDemo(); return; }
  await renderCuenta();
}

async function renderCuenta() {
  body().innerHTML = '<p class="hint">Cargando…</p>';
  $('#cuenta-sub').textContent = '';
  let sb;
  try { sb = await nube(); } catch { return sinConexion(); }
  const { data: { session } } = await sb.auth.getSession();
  perfil = null;
  if (!session) { guardarAcceso(null); guardarPaquete(null); cambiarRol(null); return renderEntrar(); }
  const { data, error } = await sb.rpc('mi_perfil');
  if (error) {
    if (/fetch|network/i.test(error.message || '')) return sinConexion(session.user.email);
    body().innerHTML = `<div class="cuenta-card"><b>No se pudo cargar tu cuenta</b><small>${escapeHtml(errorLegible(error))}</small></div>${btnSalir()}`;
    return;
  }
  const p = data && data[0];
  guardarAcceso(p);
  guardarPaquete(p ? p.paquete : null);
  cambiarRol(p ? p.rol : null);
  if (!p) {
    $('#cuenta-sub').textContent = session.user.email;
    body().innerHTML = `<div class="cuenta-card"><b>Tu cuenta aún no pertenece a una escuela</b>
      <small>Abre el enlace de invitación que te mandó tu director, o pídele uno.</small></div>${btnSalir()}${btnEliminar()}`;
    return;
  }
  perfil = p;
  $('#cuenta-sub').textContent = p.correo;
  let html = `<div class="cuenta-card">
      <b>${escapeHtml(p.nombre)}</b><small>${escapeHtml(p.correo)}</small>
      <span class="rol-badge">${ROL[p.rol]}</span>
      ${p.escuela ? `<div class="cuenta-escuela">${escapeHtml(p.escuela)} · ${escapeHtml(p.clave)}</div>` : ''}
    </div>
    ${avisoSuscripcion(p)}
    ${p.escuela_id ? '<div class="nube-estado" data-nube></div>' : ''}
    ${p.escuela_id && !(state && state.groups && state.groups.length) ? '<div class="cuenta-bajar" data-bajar hidden></div>' : ''}`;
  if (p.rol === 'director') html += btnsDirector();
  if (p.rol === 'superadmin') html += '<a class="btn primary dir-link" href="/director" target="_blank" rel="noopener">Abrir panel de escuelas ↗</a>';
  if (p.rol === 'superadmin') html += seccionVerComo() + seccionEscuelas();
  body().innerHTML = html + btnSalir() + (p.rol === 'superadmin' ? '' : btnEliminar());
  pintarNube();
  programar(200);
  ofrecerBajar();
  if (p.rol === 'superadmin') await cargarEscuelas(sb);
}

// Estado de la copia en la nube de lo capturado en este celular.
function textoNube() {
  const grupos = (state && state.groups) ? state.groups.length : 0;
  if (!grupos) return ['', 'Aún no hay listas en este celular. Cuando subas tu Excel se guardarán también en tu cuenta.'];
  const hace = estado.ultimo ? new Date(estado.ultimo).toLocaleTimeString('es-MX', { hour: 'numeric', minute: '2-digit' }) : '';
  switch (estado.fase) {
    case 'subiendo': return ['subiendo', `Guardando en tu cuenta…${estado.subidos ? ` (${estado.subidos} días)` : ''}`];
    case 'sin-conexion': return ['espera', 'Sin internet: lo capturado está seguro en este celular y se guardará en tu cuenta al volver la conexión.'];
    case 'error': return ['mal', `No se pudo guardar en tu cuenta: ${estado.error}`];
    case 'otra-cuenta': return ['mal', 'Las listas de este celular son de otra cuenta, así que no se suben a esta.'];
    case 'sin-escuela': return ['espera', 'Tu cuenta aún no pertenece a una escuela.'];
    case 'ejemplo': return ['espera', 'Estás usando la lista de ejemplo: no se guarda en tu cuenta. Cuando quieras, sube tu Excel con ⋮ → Cargar otro archivo.'];
    case 'vencida': return ['', '']; // ya lo dice el aviso de la suscripción
    case 'listo': return ['ok', `Todo guardado en tu cuenta ✓${hace ? ` · ${hace}` : ''}`];
    default: return ['', 'Tus listas se guardan en este celular y en tu cuenta.'];
  }
}
function pintarNube() {
  const el = body().querySelector('[data-nube]');
  if (!el) return;
  const [cls, txt] = textoNube();
  el.hidden = !txt;
  el.className = `nube-estado ${cls}`;
  el.innerHTML = `<span>${escapeHtml(txt)}</span>${['mal', 'espera', 'ok'].includes(cls) && estado.fase !== 'otra-cuenta' ? '<button class="link-btn" data-act="subir">Guardar ahora</button>' : ''}`;
}
alCambiarEstado(() => { if ($('#cuenta').open) pintarNube(); });

// Dispositivo sin listas: si la cuenta tiene clases guardadas, ofrecer bajarlas.
async function ofrecerBajar() {
  const box = body().querySelector('[data-bajar]');
  if (!box) return;
  let n = 0;
  try { n = await clasesEnLaNube(); } catch { return; }
  if (!n) return;
  box.hidden = false;
  const aviso = body().querySelector('[data-nube]');
  if (aviso) aviso.hidden = true;
  box.innerHTML = `<b>Tienes ${n} ${n === 1 ? 'clase guardada' : 'clases guardadas'} en tu cuenta</b>
    <small>Bájalas a este celular para seguir pasando lista y calificando aquí. Necesitas internet solo para bajarlas.</small>
    <button class="btn primary" data-act="bajar">Bajar a este celular</button><p class="hint" data-avance></p>`;
}

// Prueba gratis, suscripción vencida o escuela desactivada.
export const suscripcionVencida = (p) => !!p && !!p.escuela_id && (!p.activa || (p.vence_el && p.vence_el < new Date().toISOString().slice(0, 10)));
function avisoSuscripcion(p) {
  if (!p.escuela_id) return '';
  if (suscripcionVencida(p)) {
    return `<div class="nube-estado mal"><span>${!p.activa ? 'Tu escuela está desactivada.' : `Tu ${p.paquete === 'individual' ? 'suscripción' : 'licencia'} venció el ${p.vence_el}.`}
      Tus listas siguen guardándose en este celular, pero ya no se guardan en tu cuenta. ${p.paquete === 'individual' ? 'Renueva para seguir.' : 'Avisa a tu director.'}</span></div>`;
  }
  if (p.paquete === 'individual' && p.vence_el) {
    const dias = Math.ceil((new Date(`${p.vence_el}T23:59:59`) - new Date()) / 864e5);
    if (dias <= 31) return `<div class="nube-estado espera"><span>Tu cuenta está activa hasta el ${p.vence_el} (${dias} día${dias === 1 ? '' : 's'}).</span></div>`;
  }
  return '';
}

// Director: solo el botón para ir a la otra pantalla (si ya está en su escuela, no se le ofrece ir a ella).
const btnsDirector = () => (vistaEscuela()
  ? `<section class="cuenta-sec"><h3>¿También das clases?</h3>
      <p class="hint">Si pasas lista a algún grupo, aquí tienes tus grupos como cualquier maestro.</p>
      <button class="btn" data-act="ir-grupos">Ir a mis grupos</button></section>`
  : `<button class="btn primary big" data-act="ir-escuela">‹ Volver a mi escuela</button>
    <p class="hint">Resumen de la escuela, maestros (invitar y dar o quitar acceso) y parciales.</p>`);

// Rol guardado en el dispositivo: si cambia (entrar, salir, otra cuenta), la app cambia de pantalla.
function cambiarRol(r) {
  if (rolGuardado() === r) return;
  guardarRol(r);
  render();
}

const btnSalir = () => '<button class="btn cuenta-salir" data-act="salir">Cerrar sesión</button>';
const btnEliminar = () => `<button class="btn cuenta-eliminar" data-act="eliminar">Eliminar mi cuenta</button>
  <p class="hint cuenta-legal"><a href="/privacidad" target="_blank" rel="noopener">Aviso de privacidad</a> · <a href="/terminos" target="_blank" rel="noopener">Términos</a></p>`;

// Eliminar mi cuenta (lo piden Apple y Google). El independiente borra todo lo suyo de la nube; el de
// una escuela borra su usuario y las calificaciones se quedan con la escuela. El celular conserva sus listas.
async function eliminarCuenta(b) {
  const indep = !perfil || perfil.paquete === 'individual';
  const texto = indep
    ? 'Se borran de la nube tu cuenta, tus grupos, alumnos, calificaciones y archivos. Las listas que están en este celular se quedan aquí.'
    : `Se borran tu nombre, correo y contraseña. Las listas y calificaciones que capturaste son registros de ${perfil.escuela} y se quedan con la escuela. Lo que está en este celular se queda aquí.`;
  if (!(await confirmBulk('¿Eliminar tu cuenta?', texto, 'Sí, eliminar', true))) return;
  if (!(await confirmBulk('¿Seguro?', 'Esto no se puede deshacer.', 'Eliminar para siempre', true))) return;
  b.disabled = true;
  try {
    const sb = await nube();
    const { data: { session } } = await sb.auth.getSession();
    if (indep && perfil && perfil.escuela_id && session) await borrarArchivos(sb, `${perfil.escuela_id}/${session.user.id}`);
    const { error } = await sb.rpc('eliminar_mi_cuenta');
    if (error) throw error;
    desvincular();
    await sb.auth.signOut({ scope: 'local' });
    $('#cuenta').close();
    toast('Tu cuenta se eliminó');
  } catch (ex) {
    console.error(ex);
    toast(errorLegible(ex));
    b.disabled = false;
  }
}

async function borrarArchivos(sb, carpeta) {
  const st = sb.storage.from('archivos');
  for (const sub of ['excel', 'cuadernillos']) {
    const { data } = await st.list(`${carpeta}/${sub}`, { limit: 1000 });
    const rutas = (data || []).map((f) => `${carpeta}/${sub}/${f.name}`);
    if (rutas.length) await st.remove(rutas);
  }
}

// Las listas del celular dejan de apuntar a la cuenta borrada (para poder usar otra después).
function desvincular() {
  if (!state) return;
  const limpiar = (o) => {
    if (Array.isArray(o)) { o.forEach(limpiar); return; }
    if (!o || typeof o !== 'object') return;
    for (const k of ['nube', 'nube_id', 'alumno_id', 'nube_reabierto']) delete o[k];
    Object.values(o).forEach(limpiar);
  };
  limpiar(state.groups);
  delete state.nube_usuario;
  delete state.nube_excel;
  delete state.nube_ajustes;
  delete state.nube_ajustes_en;
  // Los parciales que puso la escuela pasan a ser del maestro (ya puede editarlos).
  if (state.periodos && state.periodos.origen === 'escuela') { state.periodos.origen = 'maestro'; delete state.periodos.id; }
  save(true);
}

function sinConexion(correo) {
  body().innerHTML = `<div class="cuenta-card"><b>${correo ? 'Sesión iniciada' : 'Sin conexión'}</b>
    <small>${correo ? `${escapeHtml(correo)} · ` : ''}Para ver tu cuenta necesitas internet. La app sigue funcionando sin conexión.</small></div>`;
}

function renderEntrar() {
  body().innerHTML = `<form class="cuenta-form" data-form="entrar" autocomplete="on">
      <p class="hint">Entra con el correo y la contraseña de tu cuenta. Si aún no tienes, pide a tu director una invitación.</p>
      <label class="field">Correo<input name="correo" type="email" autocomplete="username" required></label>
      <label class="field">Contraseña<input name="pass" type="password" autocomplete="current-password" required></label>
      <label class="check"><input type="checkbox" name="recordar" ${recordar() ? 'checked' : ''}> Mantener la sesión iniciada en este celular</label>
      <p class="form-error" data-error></p>
      <button class="btn primary big" type="submit">Entrar</button>
    </form>
    <div class="cuenta-card cuenta-indep">
      <b>¿Eres maestro independiente?</b>
      <small>Crea tu cuenta sin escuela: 30 días gratis para guardar tus listas en la nube y abrirlas en otro celular.</small>
      <button class="btn" data-act="registro">Crear mi cuenta</button>
    </div>`;
}

function renderRegistro() {
  body().innerHTML = `<form class="cuenta-form" data-form="registro" autocomplete="on">
      <p class="hint">Cuenta de maestro independiente. Si tu escuela usa la app, mejor pide a tu director una invitación.</p>
      <label class="field">Tu nombre completo<input name="nombre" autocomplete="name" required></label>
      <label class="field">Correo<input name="correo" type="email" autocomplete="username" required></label>
      <label class="field">Crea una contraseña (mínimo 8)<input name="pass" type="password" autocomplete="new-password" minlength="8" required></label>
      <label class="check"><input type="checkbox" name="recordar" checked> Mantener la sesión iniciada en este celular</label>
      <p class="form-error" data-error></p>
      <p class="hint cuenta-legal">Al crear tu cuenta aceptas los <a href="/terminos" target="_blank" rel="noopener">términos</a> y el <a href="/privacidad" target="_blank" rel="noopener">aviso de privacidad</a>.</p>
      <button class="btn primary big" type="submit">Crear cuenta (30 días gratis)</button>
      <button class="btn" type="button" data-act="volver-entrar">Ya tengo cuenta: entrar</button>
    </form>`;
}

// ---------- Director: maestros e invitaciones ----------
// ---------- Superadmin: abrir la app como cada rol (demostración en otra pestaña) ----------
function seccionVerComo() {
  const ir = (rol, txt) => `<a class="btn" href="/?demo=${rol}" target="_blank" rel="noopener">${txt} ↗</a>`;
  return `<section class="cuenta-sec">
      <h3>Abrir la app como…</h3>
      <p class="hint">Se abre en otra pestaña tal como la ve alguien que entra con ese rol, con datos de ejemplo. Lo que hagas ahí se borra al cerrarla y no toca tus datos.</p>
      <div class="ver-como">${ir('director', 'Director')}${ir('maestro', 'Maestro de escuela')}${ir('independiente', 'Maestro independiente')}</div>
    </section>`;
}

// «Mi cuenta» en la demostración: el perfil de ejemplo de ese rol.
const enDias = (n) => new Date(Date.now() + n * 864e5).toISOString().slice(0, 10);
function renderDemo() {
  const p = { ...PERFIL_DEMO[DEMO], escuela_id: 'demo', activa: true, vence_el: DEMO === 'independiente' ? enDias(12) : null };
  $('#cuenta-sub').textContent = p.correo;
  const listas = !!(state && state.groups && state.groups.length);
  body().innerHTML = `<div class="cuenta-demo">
    <div class="cuenta-card">
      <b>${escapeHtml(p.nombre)}</b><small>${escapeHtml(p.correo)}</small>
      <span class="rol-badge">${ROL[p.rol]}</span>
      <div class="cuenta-escuela">${escapeHtml(p.escuela)} · ${escapeHtml(p.clave)}</div>
    </div>
    ${avisoSuscripcion(p)}
    <div class="nube-estado ${listas ? 'ok' : ''}"><span>${listas ? 'Todo guardado en tu cuenta ✓' : 'Aún no hay listas en este celular. Cuando subas tu Excel se guardarán también en tu cuenta.'}</span></div>
    ${p.rol === 'director' ? btnsDirector() : ''}
    <button class="btn primary cuenta-salir" data-act="salir">‹ Salir de la demostración</button>${btnEliminar()}</div>`;
}

// ---------- Superadmin: escuelas ----------
function seccionEscuelas() {
  return `<section class="cuenta-sec">
      <h3>Escuelas</h3>
      <form class="cuenta-form" data-form="escuela">
        <label class="field">Nombre de la escuela<input name="nombre" required></label>
        <div class="cuenta-2col">
          <label class="field">Clave corta<input name="clave" required maxlength="10" placeholder="ESC01" class="mayus"></label>
          <label class="field">Cuentas<input name="cuentas" type="number" min="1" value="5" required></label>
        </div>
        <div class="cuenta-2col"><label class="field">Correo de quien invitas<input name="correo" type="email" required></label>
          <label class="field">Como<select name="rol"><option value="director">Director</option><option value="maestro">Maestro</option></select></label></div>
        <button class="btn primary" type="submit">Dar de alta e invitar</button>
        <p class="form-error" data-error></p>
      </form>
      <div data-nuevo></div>
      <ul class="pick-list cuenta-lista" data-escuelas></ul>
      <h3>Precios de lista (MXN por maestro)</h3>
      <div data-precios></div>
    </section>`;
}

// Panel: cada escuela con su paquete, cuentas, precio (de lista o especial), vencimiento y pagos.
const pesos = (n) => `$${Number(n || 0).toLocaleString('es-MX', { maximumFractionDigits: 2 })}`;
const hoyISO = () => new Date().toISOString().slice(0, 10);
async function cargarEscuelas(sb) {
  const [{ data, error }, precios] = await Promise.all([sb.rpc('panel_escuelas'), sb.from('precios_lista').select('*').order('paquete')]);
  const ul = body().querySelector('[data-escuelas]');
  if (error) { ul.innerHTML = `<li class="empty">${escapeHtml(errorLegible(error))}</li>`; return; }
  const activas = (data || []).filter((e) => e.activa && (!e.vence_el || e.vence_el >= hoyISO()));
  const mensual = activas.reduce((n, e) => n + e.cuentas_contratadas * Number(e.precio_mensual ?? e.precio_lista ?? 0), 0);
  ul.innerHTML = (data || []).length
    ? `<li class="titulo">Escuelas: ${data.length} · activas ${activas.length} · ${pesos(mensual)} al mes</li>` + data.map((e) => {
        const precio = e.precio_mensual ?? e.precio_lista;
        const vencida = e.vence_el && e.vence_el < hoyISO();
        const estado = !e.activa ? '<b class="danger-text">inactiva</b>' : vencida ? `<b class="danger-text">venció ${e.vence_el}</b>` : e.vence_el ? `vence ${e.vence_el}` : 'sin vencimiento';
        return `<li class="esc-item" data-esc="${e.id}">
          <details>
            <summary><span class="nm">${escapeHtml(e.nombre)}<small>${escapeHtml(e.clave)} · ${e.paquete} · ${e.cuentas_usadas}/${e.cuentas_contratadas} cuentas
              · ${pesos(precio)}${e.precio_mensual != null ? ' (especial)' : ''} por maestro · ${estado}</small></span></summary>
            <form class="cuenta-form esc-form" data-form="esc-editar">
              <div class="cuenta-2col"><label class="field">Paquete<select name="paquete"><option value="empresarial" ${e.paquete === 'empresarial' ? 'selected' : ''}>Empresarial</option><option value="individual" ${e.paquete === 'individual' ? 'selected' : ''}>Individual</option></select></label>
                <label class="field">Cuentas<input name="cuentas" type="number" min="1" value="${e.cuentas_contratadas}"></label></div>
              <div class="cuenta-2col"><label class="field">Precio especial por maestro (vacío = lista ${pesos(e.precio_lista)})<input name="precio" inputmode="decimal" value="${e.precio_mensual ?? ''}"></label>
                <label class="field">Vence<input name="vence" type="date" value="${e.vence_el || ''}"></label></div>
              <label class="check"><input type="checkbox" name="activa" ${e.activa ? 'checked' : ''}> Activa</label>
              <button class="btn primary" type="submit">Guardar cambios</button>
              <p class="form-error" data-error></p>
            </form>
            <form class="cuenta-form esc-form" data-form="esc-invitar">
              <b>Invitar a esta escuela</b>
              <div class="cuenta-2col"><label class="field">Correo<input name="correo" type="email" required></label>
                <label class="field">Como<select name="rol"><option value="maestro">Maestro</option><option value="director">Director</option></select></label></div>
              <button class="btn" type="submit">Crear invitación</button>
              <p class="form-error" data-error></p>
              <div data-enlace></div>
            </form>
            <form class="cuenta-form esc-form" data-form="esc-pago">
              <b>Registrar pago</b>
              <div class="cuenta-2col"><label class="field">Monto (MXN)<input name="monto" inputmode="decimal" required value="${e.cuentas_contratadas * Number(precio || 0) * 6}"></label>
                <label class="field">Método<select name="metodo"><option>transferencia</option><option>efectivo</option><option>tarjeta</option><option>depósito</option></select></label></div>
              <div class="cuenta-2col"><label class="field">Del<input name="inicio" type="date" required value="${hoyISO()}"></label>
                <label class="field">Al<input name="fin" type="date" required></label></div>
              <label class="field">Factura (folio, opcional)<input name="factura"></label>
              <button class="btn" type="submit">Registrar pago y extender</button>
              <p class="form-error" data-error></p>
            </form>
          </details></li>`;
      }).join('')
    : '<li class="empty">Aún no hay escuelas.</li>';
  const pl = body().querySelector('[data-precios]');
  if (pl && precios.data) {
    pl.innerHTML = precios.data.map((p) => `<form class="cuenta-form precio-form" data-form="precio" data-paquete="${p.paquete}">
        <b>${p.paquete === 'individual' ? 'Individual' : 'Empresarial'}</b>
        <label>Mes <input name="mes" inputmode="decimal" value="${Number(p.por_mes)}"></label>
        <label>Año <input name="anio" inputmode="decimal" value="${Number(p.por_anio)}"></label>
        <label>Prueba <input name="prueba" type="number" min="0" value="${p.dias_prueba}"> días</label>
        <button class="btn" type="submit">Guardar</button><p class="form-error" data-error></p>
      </form>`).join('');
  }
}

// Muestra el enlace de una invitación recién creada, listo para mandarlo por WhatsApp o correo.
function mostrarEnlace(token, correo, box = body().querySelector('[data-nuevo]')) {
  box.innerHTML = `<div class="cuenta-enlace" data-token="${token}">
      <b>Invitación lista para ${escapeHtml(correo)}</b>
      <small>Mándale este enlace. Sirve una sola vez, solo con ese correo, y vence en 7 días.</small>
      <input readonly value="${enlace(token)}">
      <div class="row"><button type="button" class="btn" data-act="copiar">Copiar</button><button type="button" class="btn primary" data-act="compartir">Enviar</button></div>
    </div>`;
}

async function compartir(token) {
  const url = enlace(token);
  const text = 'Te invito a la app para pasar lista y calificar de nuestra escuela. Crea tu cuenta aquí:';
  if (navigator.share) {
    try { await navigator.share({ title: 'Invitación', text, url }); return; } catch { /* cancelado: se copia */ }
  }
  await copiar(url);
}
async function copiar(texto) {
  try { await navigator.clipboard.writeText(texto); toast('Enlace copiado'); } catch { toast('Copia el enlace manualmente'); }
}

// ---------- Eventos ----------
$('#cuenta-body').addEventListener('submit', async (e) => {
  e.preventDefault();
  const f = e.target;
  const err = f.querySelector('[data-error]');
  const btn = f.querySelector('[type=submit]');
  err.textContent = '';
  btn.disabled = true;
  try {
    const sb = await nube();
    const v = (n) => (f.elements[n] ? f.elements[n].value.trim() : '');
    if (f.dataset.form === 'entrar') {
      setRecordar(f.elements.recordar.checked);
      const { error } = await sb.auth.signInWithPassword({ email: v('correo').toLowerCase(), password: f.elements.pass.value });
      if (error) throw error;
      toast('Sesión iniciada');
      await renderCuenta();
    } else if (f.dataset.form === 'registro') {
      setRecordar(f.elements.recordar.checked);
      const email = v('correo').toLowerCase();
      const password = f.elements.pass.value;
      let r = await sb.auth.signUp({ email, password, options: { data: { nombre: v('nombre') } } });
      if (r.error && /already registered|already been registered/i.test(r.error.message)) r = await sb.auth.signInWithPassword({ email, password });
      if (r.error) throw r.error;
      if (!r.data.session) throw new Error('Falta confirmar el correo. Avisa al administrador.');
      const { error } = await sb.rpc('registrar_individual', { p_nombre: v('nombre') });
      if (error && !/ya pertenece/.test(error.message)) throw error;
      toast('¡Listo! Tu cuenta quedó creada');
      await renderCuenta();
      programar(300);
    } else if (f.dataset.form === 'esc-invitar') {
      const correo = v('correo').toLowerCase();
      const { data, error } = await sb.from('invitaciones')
        .insert({ escuela_id: f.closest('[data-esc]').dataset.esc, correo, rol: v('rol') === 'director' ? 'director' : 'maestro' }).select('token').single();
      if (error) throw error;
      f.elements.correo.value = '';
      mostrarEnlace(data.token, correo, f.querySelector('[data-enlace]'));
    } else if (f.dataset.form === 'esc-editar' || f.dataset.form === 'esc-pago') {
      const id = f.closest('[data-esc]').dataset.esc;
      if (f.dataset.form === 'esc-editar') {
        const precio = v('precio') === '' ? null : Number(v('precio').replace(',', '.'));
        if (precio != null && !(precio >= 0)) throw new Error('El precio debe ser un número.');
        const { error } = await sb.from('escuelas').update({
          paquete: v('paquete'), cuentas_contratadas: Math.max(1, Number(v('cuentas')) || 1), precio_mensual: precio,
          vence_el: v('vence') || null, activa: f.elements.activa.checked,
        }).eq('id', id);
        if (error) throw error;
        toast('Escuela actualizada');
      } else {
        if (v('fin') < v('inicio')) throw new Error('La fecha final debe ser después de la inicial.');
        const { error } = await sb.rpc('registrar_pago', { p_escuela: id, p_monto: Number(v('monto').replace(',', '.')) || 0, p_inicio: v('inicio'), p_fin: v('fin'), p_metodo: v('metodo'), p_factura: v('factura') });
        if (error) throw error;
        toast(`Pago registrado: la escuela queda activa hasta ${v('fin')}`);
      }
      await cargarEscuelas(sb);
    } else if (f.dataset.form === 'precio') {
      const { error } = await sb.from('precios_lista').update({ por_mes: Number(v('mes')) || 0, por_anio: Number(v('anio')) || 0, dias_prueba: Number(v('prueba')) || 0, actualizado_en: new Date().toISOString() }).eq('paquete', f.dataset.paquete);
      if (error) throw error;
      toast('Precio de lista guardado');
      await cargarEscuelas(sb);
    } else if (f.dataset.form === 'escuela') {
      const correo = v('correo').toLowerCase();
      const clave = v('clave').toUpperCase();
      const { data: esc, error } = await sb.from('escuelas')
        .insert({ nombre: v('nombre'), clave, cuentas_contratadas: Number(v('cuentas')) || 1, paquete: 'empresarial' }).select('id').single();
      if (error) throw error;
      const inv = await sb.from('invitaciones').insert({ escuela_id: esc.id, correo, rol: v('rol') === 'maestro' ? 'maestro' : 'director' }).select('token').single();
      if (inv.error) throw inv.error;
      f.reset();
      await cargarEscuelas(sb);
      mostrarEnlace(inv.data.token, correo);
    }
  } catch (ex) {
    console.error(ex);
    const m = String(ex.message || '');
    err.textContent = /cuentas contratadas/.test(m) ? 'Ya usaste todas las cuentas contratadas. Pide más para invitar a otro maestro.'
      : /escuelas_clave_key|duplicate key/.test(m) ? 'Ya existe una escuela con esa clave.'
      : /clave_check/.test(m) ? 'La clave lleva solo letras y números (2 a 10), por ejemplo ESC01.'
      : errorLegible(ex);
  } finally {
    btn.disabled = false;
  }
});

$('#cuenta-body').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-act]');
  if (!b) return;
  const act = b.dataset.act;
  if (act === 'ir-escuela') { setVistaApp('escuela'); $('#cuenta').close(); render(); return; }
  if (act === 'ir-grupos') { setVistaApp('grupos'); $('#cuenta').close(); render(); return; }
  if (DEMO) {
    // Cerrar sesión en la demostración = volver a tu propia app (con tu cuenta y tus datos).
    if (act === 'salir') { location.href = '/'; return; }
    toast('Demostración: aquí no se guarda nada');
    return;
  }
  const token = (b.closest('[data-token]') || {}).dataset?.token;
  if (act === 'subir') sincronizar();
  if (act === 'registro') renderRegistro();
  if (act === 'eliminar') eliminarCuenta(b);
  if (act === 'volver-entrar') renderEntrar();
  if (act === 'bajar') {
    const av = body().querySelector('[data-avance]');
    b.disabled = true;
    try {
      const n = await bajarDeLaNube((t) => { if (av) av.textContent = t; });
      $('#cuenta').close();
      render();
      toast(`Listo: ${n} ${n === 1 ? 'clase bajada' : 'clases bajadas'} a este celular`);
      programar(500); // trae también los ajustes de la cuenta (valores de cada estado, parciales)
    } catch (ex) {
      console.error(ex);
      if (av) av.textContent = errorLegible(ex);
      b.disabled = false;
    }
  }
  if (act === 'copiar') copiar(enlace(token));
  if (act === 'compartir') compartir(token);
  if (act === 'borrar-inv') {
    if (!(await confirmBulk('¿Quitar esta invitación?', 'El enlace dejará de funcionar y se libera la cuenta.', 'Sí, quitar', true))) return;
    const sb = await nube();
    const { error } = await sb.from('invitaciones').delete().eq('id', b.dataset.id);
    if (error) { toast(errorLegible(error)); return; }
    await renderCuenta();
  }
  if (act === 'salir') {
    if (!(await confirmBulk('¿Cerrar sesión?', 'Tus listas siguen guardadas en este celular.', 'Sí, cerrar sesión', true))) return;
    const sb = await nube();
    await sb.auth.signOut();
    toast('Sesión cerrada');
    await renderCuenta();
  }
});

// ---------- Aceptar una invitación (enlace ?invitacion=TOKEN) ----------
let invToken = null;
export async function openInvitacion(token) {
  invToken = token;
  const d = $('#invitacion');
  $('#inv-titulo').textContent = 'Invitación';
  $('#inv-texto').textContent = 'Revisando la invitación…';
  $('#inv-form').hidden = true;
  $('#inv-cerrar').hidden = false;
  d.showModal();
  try {
    const sb = await nube();
    const { data, error } = await sb.rpc('ver_invitacion', { p_token: token });
    if (error) throw error;
    const inv = data && data[0];
    if (!inv || !inv.valida) {
      $('#inv-texto').textContent = 'Esta invitación ya no es válida (se usó, venció o se quitó). Pide otra a tu director.';
      return;
    }
    $('#inv-titulo').textContent = `Te invitaron a ${inv.escuela}`;
    $('#inv-texto').textContent = `Como ${ROL[inv.rol].toLowerCase()}. Crea tu cuenta para entrar con tu correo y contraseña. Si ya tienes cuenta, escribe tu contraseña de siempre.`;
    $('#inv-correo').value = inv.correo;
    $('#inv-recordar').checked = recordar();
    $('#inv-error').textContent = '';
    $('#inv-form').hidden = false;
    $('#inv-cerrar').hidden = true;
    $('#inv-nombre').focus();
  } catch (ex) {
    $('#inv-texto').textContent = errorLegible(ex);
  }
}
$('#invitacion').addEventListener('close', () => { if (location.search.includes('invitacion=')) history.replaceState(null, '', location.pathname); });
$('#inv-form').addEventListener('submit', async (e) => {
  e.preventDefault();
  const btn = $('#inv-ok');
  const err = $('#inv-error');
  err.textContent = '';
  btn.disabled = true;
  const email = $('#inv-correo').value;
  const password = $('#inv-pass').value;
  const nombre = $('#inv-nombre').value.trim();
  try {
    const sb = await nube();
    setRecordar($('#inv-recordar').checked);
    let { data, error } = await sb.auth.signUp({ email, password, options: { data: { nombre } } });
    if (error && /already registered|already been registered/i.test(error.message)) {
      ({ data, error } = await sb.auth.signInWithPassword({ email, password }));
      if (error) throw new Error('Ese correo ya tiene cuenta y la contraseña no coincide.');
    } else if (error) throw error;
    if (!data.session) throw new Error('Falta confirmar el correo. Avisa al administrador.');
    const acepta = await sb.rpc('aceptar_invitacion', { p_token: invToken, p_nombre: nombre });
    if (acepta.error) throw acepta.error;
    history.replaceState(null, '', location.pathname);
    $('#invitacion').close();
    toast('¡Listo! Tu cuenta quedó creada');
    // El director va directo a su escuela; el maestro ve su cuenta (para subir su Excel o bajar sus listas).
    const rol = acepta.data && (acepta.data.rol || (acepta.data[0] && acepta.data[0].rol));
    if (rol === 'director') { guardarRol('director'); setVistaApp('escuela'); render(); } else openCuenta();
    programar(300);
  } catch (ex) {
    console.error(ex);
    err.textContent = /ya pertenece/.test(ex.message) ? 'Esta cuenta ya pertenece a una escuela.' : errorLegible(ex);
  } finally {
    btn.disabled = false;
  }
});

