// ---------- Panel del director (/director): pensado para computadora ----------
// Entrar, elegir escuela (administrador) y mostrar el panel (src/director/panel.js).
import { errorLegible, nube } from '../lib/nube.js';
import { crearPanel, escuelaDeEjemplo, exportarEscuela, leerEscuela } from './panel.js';

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const panel = crearPanel($('#dir-main'));
let sb;
let actual = null; // [escuela, nombre]

// Escuela de ejemplo (?ejemplo=1): datos inventados, sin entrar a ninguna cuenta.
const EJEMPLO = new URLSearchParams(location.search).has('ejemplo');
async function iniciarEjemplo() {
  document.body.insertAdjacentHTML('afterbegin', '<div class="dir-ejemplo">Escuela de ejemplo · datos inventados para conocer el panel</div>');
  $('#dir-sub').textContent = 'Preparatoria de Ejemplo · EJEMPLO';
  panel.mostrar(await escuelaDeEjemplo());
  $('#dir-exportar').hidden = false;
}

async function iniciar() {
  if (EJEMPLO) return iniciarEjemplo();
  try { sb = await nube(); } catch { $('#dir-main').innerHTML = '<p class="hint">Sin conexión. El panel necesita internet.</p>'; return; }
  const { data: { session } } = await sb.auth.getSession();
  if (!session) return pedirEntrar();
  const { data: p, error } = await sb.rpc('mi_perfil');
  if (error) { $('#dir-main').innerHTML = `<p class="hint">${esc(errorLegible(error))}</p>`; return; }
  const perfil = p && p[0];
  if (!perfil || !['director', 'superadmin'].includes(perfil.rol)) {
    $('#dir-main').innerHTML = '<div class="cuenta-card"><b>Este panel es para directores</b><small>Tu cuenta no tiene acceso. Regresa a la app para pasar lista.</small></div>';
    $('#dir-sub').textContent = perfil ? perfil.correo : '';
    return;
  }
  if (perfil.rol === 'superadmin') {
    const { data: escuelas } = await sb.rpc('panel_escuelas');
    const sel = $('#dir-escuela');
    sel.innerHTML = (escuelas || []).map((e) => `<option value="${e.id}">${esc(e.nombre)} · ${esc(e.clave)}</option>`).join('');
    sel.hidden = !(escuelas || []).length;
    sel.onchange = () => cargar(sel.value, sel.selectedOptions[0].textContent);
    if (!(escuelas || []).length) { $('#dir-main').innerHTML = '<p class="hint">Aún no hay escuelas.</p>'; return; }
    return cargar(sel.value, sel.selectedOptions[0].textContent);
  }
  return cargar(perfil.escuela_id, `${perfil.escuela} · ${perfil.clave}`);
}

function pedirEntrar() {
  $('#dir-sub').textContent = 'Entra con tu cuenta de director';
  $('#dir-main').innerHTML = `<form class="cuenta-form dir-login" id="dir-login">
      <label class="field">Correo<input name="correo" type="email" autocomplete="username" required></label>
      <label class="field">Contraseña<input name="pass" type="password" autocomplete="current-password" required></label>
      <p class="form-error" id="dir-login-err"></p>
      <button class="btn primary big" type="submit">Entrar</button>
    </form>`;
  $('#dir-login').onsubmit = async (e) => {
    e.preventDefault();
    const f = e.target;
    const { error } = await sb.auth.signInWithPassword({ email: f.correo.value.trim().toLowerCase(), password: f.pass.value });
    if (error) { $('#dir-login-err').textContent = errorLegible(error); return; }
    iniciar();
  };
}


async function cargar(escuela, nombre) {
  actual = [escuela, nombre];
  $('#dir-sub').textContent = nombre;
  $('#dir-main').innerHTML = '<p class="hint">Cargando datos de la escuela…</p>';
  try {
    panel.mostrar(await leerEscuela(sb, escuela));
    $('#dir-recargar').hidden = false;
    $('#dir-exportar').hidden = false;
  } catch (e) {
    console.error(e);
    $('#dir-main').innerHTML = `<p class="hint">No se pudo cargar: ${esc(errorLegible(e))}</p>`;
  }
}

$('#dir-recargar').addEventListener('click', () => actual && cargar(...actual));
$('#dir-exportar').addEventListener('click', async (e) => {
  e.currentTarget.disabled = true;
  try { await exportarEscuela(panel.datos(), $('#dir-sub').textContent.split(' · ')[0]); } catch (ex) { console.error(ex); alertar('No se pudo exportar. Revisa tu conexión.'); }
  $('#dir-exportar').disabled = false;
});
const alertar = (t) => { $('#dir-main').insertAdjacentHTML('afterbegin', `<p class="form-error">${esc(t)}</p>`); };

iniciar();
