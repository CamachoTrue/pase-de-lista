// ---------- Mi cuenta (Supabase) ----------
// Entrar con correo y contraseña, mantener la sesión, invitar maestros (director) y dar de alta
// escuelas (superadmin). Por ahora las listas y calificaciones siguen guardándose en el celular;
// la sincronización con la cuenta llega en la etapa 4.
import { $, escapeHtml, toast } from './util.js';
import { confirmBulk } from './dialogs.js';
import { errorLegible, nube, recordar, setRecordar } from '../lib/nube.js';
import { alCambiarEstado, estado, programar, sincronizar } from './sync.js';
import { state } from './state.js';
import { bajarDeLaNube, clasesEnLaNube } from './bajar.js';
import { render } from './main-view.js';

const ROL = { superadmin: 'Administrador general', director: 'Director', maestro: 'Maestro' };
const body = () => $('#cuenta-body');
const fecha = (iso) => new Date(iso).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });
const enlace = (token) => `${location.origin}/?invitacion=${token}`;

export async function openCuenta() {
  $('#cuenta').showModal();
  await renderCuenta();
}

async function renderCuenta() {
  body().innerHTML = '<p class="hint">Cargando…</p>';
  $('#cuenta-sub').textContent = '';
  let sb;
  try { sb = await nube(); } catch { return sinConexion(); }
  const { data: { session } } = await sb.auth.getSession();
  if (!session) return renderEntrar();
  const { data, error } = await sb.rpc('mi_perfil');
  if (error) {
    if (/fetch|network/i.test(error.message || '')) return sinConexion(session.user.email);
    body().innerHTML = `<div class="cuenta-card"><b>No se pudo cargar tu cuenta</b><small>${escapeHtml(errorLegible(error))}</small></div>${btnSalir()}`;
    return;
  }
  const p = data && data[0];
  if (!p) {
    $('#cuenta-sub').textContent = session.user.email;
    body().innerHTML = `<div class="cuenta-card"><b>Tu cuenta aún no pertenece a una escuela</b>
      <small>Abre el enlace de invitación que te mandó tu director, o pídele uno.</small></div>${btnSalir()}`;
    return;
  }
  $('#cuenta-sub').textContent = p.correo;
  let html = `<div class="cuenta-card">
      <b>${escapeHtml(p.nombre)}</b><small>${escapeHtml(p.correo)}</small>
      <span class="rol-badge">${ROL[p.rol]}</span>
      ${p.escuela ? `<div class="cuenta-escuela">${escapeHtml(p.escuela)} · ${escapeHtml(p.clave)}</div>` : ''}
    </div>
    ${p.escuela_id ? '<div class="nube-estado" data-nube></div>' : ''}
    ${p.escuela_id && !(state && state.groups && state.groups.length) ? '<div class="cuenta-bajar" data-bajar hidden></div>' : ''}`;
  if (p.rol === 'director') html += seccionMaestros();
  if (p.rol === 'superadmin') html += seccionEscuelas();
  body().innerHTML = html + btnSalir();
  pintarNube();
  programar(200);
  ofrecerBajar();
  if (p.rol === 'director') await cargarMaestros(sb, p.escuela_id);
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
    case 'listo': return ['ok', `Todo guardado en tu cuenta ✓${hace ? ` · ${hace}` : ''}`];
    default: return ['', 'Tus listas se guardan en este celular y en tu cuenta.'];
  }
}
function pintarNube() {
  const el = body().querySelector('[data-nube]');
  if (!el) return;
  const [cls, txt] = textoNube();
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

const btnSalir = () => '<button class="btn cuenta-salir" data-act="salir">Cerrar sesión</button>';

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
    </form>`;
}

// ---------- Director: maestros e invitaciones ----------
function seccionMaestros() {
  return `<section class="cuenta-sec">
      <h3>Maestros de tu escuela</h3>
      <p class="hint" data-cupo></p>
      <form class="cuenta-form invitar" data-form="invitar">
        <label class="field">Correo del maestro<input name="correo" type="email" autocomplete="off" required placeholder="maestro@correo.com"></label>
        <button class="btn primary" type="submit">Crear invitación</button>
        <p class="form-error" data-error></p>
      </form>
      <div data-nuevo></div>
      <ul class="pick-list cuenta-lista" data-invitaciones></ul>
      <ul class="pick-list cuenta-lista" data-usuarios></ul>
    </section>`;
}

async function cargarMaestros(sb, escuela) {
  const [cobro, inv, us] = await Promise.all([
    sb.rpc('datos_cobro'),
    sb.from('invitaciones').select('id, correo, token, vence_en, usada_en, rol').eq('escuela_id', escuela).is('usada_en', null).order('creada_en', { ascending: false }),
    sb.from('usuarios').select('id, nombre, correo, rol, activo').eq('escuela_id', escuela).order('nombre'),
  ]);
  const c = cobro.data && cobro.data[0];
  const pend = (inv.data || []).filter((i) => new Date(i.vence_en) > new Date());
  const cupo = body().querySelector('[data-cupo]');
  if (cupo && c) cupo.textContent = `${c.cuentas_usadas + pend.length} de ${c.cuentas_contratadas} cuentas en uso (incluye invitaciones pendientes).`;
  body().querySelector('[data-invitaciones]').innerHTML = pend.length
    ? '<li class="titulo">Invitaciones pendientes</li>' + pend.map((i) => `<li data-token="${i.token}">
        <span class="nm">${escapeHtml(i.correo)}<small>vence ${fecha(i.vence_en)}</small></span>
        <button class="link-btn" data-act="compartir">Enviar</button>
        <button class="link-btn danger-text" data-act="borrar-inv" data-id="${i.id}">Quitar</button></li>`).join('')
    : '';
  body().querySelector('[data-usuarios]').innerHTML = '<li class="titulo">Cuentas</li>' + (us.data || []).map((u) => `<li>
      <span class="nm">${escapeHtml(u.nombre)}<small>${escapeHtml(u.correo)} · ${ROL[u.rol]}${u.activo ? '' : ' · desactivada'}</small></span></li>`).join('');
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
        <label class="field">Correo del director<input name="correo" type="email" required></label>
        <button class="btn primary" type="submit">Dar de alta e invitar al director</button>
        <p class="form-error" data-error></p>
      </form>
      <div data-nuevo></div>
      <ul class="pick-list cuenta-lista" data-escuelas></ul>
    </section>`;
}

async function cargarEscuelas(sb) {
  const { data } = await sb.from('escuelas').select('id, nombre, clave, paquete, activa').order('creado_en', { ascending: false });
  body().querySelector('[data-escuelas]').innerHTML = (data || []).length
    ? '<li class="titulo">Escuelas dadas de alta</li>' + data.map((e) => `<li><span class="nm">${escapeHtml(e.nombre)}<small>${escapeHtml(e.clave)} · ${e.paquete}${e.activa ? '' : ' · inactiva'}</small></span></li>`).join('')
    : '<li class="empty">Aún no hay escuelas.</li>';
}

// Muestra el enlace de una invitación recién creada, listo para mandarlo por WhatsApp o correo.
function mostrarEnlace(token, correo) {
  const box = body().querySelector('[data-nuevo]');
  box.innerHTML = `<div class="cuenta-enlace" data-token="${token}">
      <b>Invitación lista para ${escapeHtml(correo)}</b>
      <small>Mándale este enlace. Sirve una sola vez, solo con ese correo, y vence en 7 días.</small>
      <input readonly value="${enlace(token)}">
      <div class="row"><button class="btn" data-act="copiar">Copiar</button><button class="btn primary" data-act="compartir">Enviar</button></div>
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
    } else if (f.dataset.form === 'invitar') {
      const { data: perfil } = await sb.rpc('mi_perfil');
      const correo = v('correo').toLowerCase();
      const { data, error } = await sb.from('invitaciones').insert({ escuela_id: perfil[0].escuela_id, correo, rol: 'maestro' }).select('token').single();
      if (error) throw error;
      f.reset();
      await cargarMaestros(sb, perfil[0].escuela_id);
      mostrarEnlace(data.token, correo);
    } else if (f.dataset.form === 'escuela') {
      const correo = v('correo').toLowerCase();
      const clave = v('clave').toUpperCase();
      const { data: esc, error } = await sb.from('escuelas')
        .insert({ nombre: v('nombre'), clave, cuentas_contratadas: Number(v('cuentas')) || 1, paquete: 'empresarial' }).select('id').single();
      if (error) throw error;
      const inv = await sb.from('invitaciones').insert({ escuela_id: esc.id, correo, rol: 'director' }).select('token').single();
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
  const token = (b.closest('[data-token]') || {}).dataset?.token;
  if (act === 'subir') sincronizar();
  if (act === 'bajar') {
    const av = body().querySelector('[data-avance]');
    b.disabled = true;
    try {
      const n = await bajarDeLaNube((t) => { if (av) av.textContent = t; });
      $('#cuenta').close();
      render();
      toast(`Listo: ${n} ${n === 1 ? 'clase bajada' : 'clases bajadas'} a este celular`);
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
    openCuenta();
    programar(300);
  } catch (ex) {
    console.error(ex);
    err.textContent = /ya pertenece/.test(ex.message) ? 'Esta cuenta ya pertenece a una escuela.' : errorLegible(ex);
  } finally {
    btn.disabled = false;
  }
});

// Pantalla inicial (sin Excel): entrar con la cuenta para bajar los datos.
$('#upload-cuenta').addEventListener('click', openCuenta);
