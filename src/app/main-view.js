import { STATUSES } from '../lib/excel-core.js';
import { renderCard, renderStudents, startCard } from './card.js';
import { renderCompact } from './compact.js';
import { OPTIONS, dayData, dayRecords, glabel, group, lista, save, secs, state, tab, todayISO, typeOf, ui, vistaApp } from './state.js';
import { haySesion, rolGuardado } from '../lib/nube.js';
import { DEMO } from '../lib/demo.js';
import { tieneAcceso } from './acceso.js';
import { accesoVencido } from '../lib/nube.js';
import { $, escapeHtml } from './util.js';

// ---------- Render ----------
// El director ve su escuela (no necesita grupos), salvo que elija "Mis grupos" porque también da clases.
export const vistaEscuela = () => rolGuardado() === 'director' && vistaApp() !== 'grupos';

export function render() {
  const escuela = vistaEscuela();
  $('#screen-escuela').hidden = !escuela;
  $('#upload-escuela').hidden = rolGuardado() !== 'director';
  // Con la sesión ya iniciada no se dice "Ya tengo cuenta: entrar".
  const dentro = haySesion() || !!DEMO;
  $('#upload-cuenta').textContent = dentro ? 'Mi cuenta' : 'Ya tengo cuenta: entrar';
  $('#upload-cuenta-hint').textContent = dentro ? 'Si ya usas la app en otro celular, en Mi cuenta puedes bajar tus grupos y calificaciones.'
    : 'Si ya usas la app en otro celular, entra para bajar tus grupos y calificaciones.';
  if (escuela) {
    $('#screen-upload').hidden = true;
    $('#screen-main').hidden = true;
    import('./escuela.js').then((m) => m.openEscuela()).catch((e) => console.error(e));
    return;
  }
  const has = !!(state && state.groups && state.groups.length);
  pintarAcceso();
  $('#screen-upload').hidden = has;
  $('#screen-main').hidden = !has;
  if (!has) return;

  const sel = $('#group-select');
  sel.innerHTML = state.groups
    .map((g, i) => `<option value="${i}">${escapeHtml(glabel(g))}</option>`)
    .join('');
  sel.value = state.ui.group;
  pintarFecha();
  startCard();
  refresh();
}


// Aviso arriba de la lista cuando no se puede capturar (sin cuenta o vencida) o es la lista de ejemplo.
function pintarAcceso() {
  const acceso = tieneAcceso();
  const ejemplo = !!(state && state.ejemplo);
  const aviso = $('#aviso-acceso');
  const txt = !acceso && haySesion() && accesoVencido() ? 'Tu acceso venció: puedes ver y exportar, pero no capturar. <b>Ver mi cuenta ›</b>'
    : ejemplo ? (acceso ? 'Estás usando la lista de ejemplo. <b>Subir mi Excel ›</b>' : 'Estás usando la lista de ejemplo. <b>Crear mi cuenta ›</b>')
    : !acceso ? 'Solo lectura: entra a tu cuenta para capturar. <b>Entrar ›</b>' : '';
  aviso.hidden = !txt;
  aviso.innerHTML = txt;
  aviso.dataset.ir = ejemplo && acceso ? 'excel' : 'cuenta';
  aviso.classList.toggle('mal', !acceso && !ejemplo);
  // Pantalla inicial sin cuenta: primero entrar; subir el Excel pide la cuenta.
  $('#upload-registro').hidden = acceso;
  $('#upload-registro').classList.toggle('primary', !acceso);
  $('#upload-subir').classList.toggle('primary', acceso);
}

export function refresh() {
  const opt = $('#group-select').options[state.ui.group];
  if (opt) opt.textContent = glabel(group());
  // Pestañas de abajo: Pasar lista, Calificaciones o Grupo.
  const vista = ui.vista || 'lista';
  document.querySelectorAll('#tabbar [data-vista]').forEach((b) => b.classList.toggle('on', b.dataset.vista === vista));
  $('#lista-controles').hidden = vista !== 'lista';
  $('#vista-calif').hidden = vista !== 'calif';
  $('#vista-grupo').hidden = vista !== 'grupo';
  if (vista !== 'lista') {
    $('#card-view').hidden = true;
    $('#list-view').hidden = true;
    $('#compact-view').hidden = true;
    const carga = vista === 'calif' ? import('./calificaciones.js').then((m) => m.renderCalif()) : import('./grupo.js').then((m) => m.renderGrupo());
    carga.catch((e) => { console.error(e); $(`#vista-${vista}`).innerHTML = '<p class="hint">No se pudo abrir. Revisa tu conexión e inténtalo otra vez.</p>'; });
    return;
  }
  const t = tab();
  renderTabs();
  const att = typeOf(t) === 'asist';
  $('#card-view').hidden = !att || ui.view !== 'card';
  $('#list-view').hidden = !att || ui.view !== 'list';
  $('#compact-view').hidden = att;
  if (!att) renderCompact();
  else if (ui.view === 'card') renderCard(false);
  else renderStudents();
  renderSummary();
}

// ---------- Pestañas (una por sección del grupo) ----------
export function renderTabs() {
  const t = tab();
  const box = $('#tabs');
  const html = secs().map((x) => `<button class="tab ${x.id === t ? 'on' : ''}" role="tab" aria-selected="${x.id === t}" data-tab="${x.id}">${escapeHtml(x.name)}</button>`).join('');
  if (box.innerHTML !== html) {
    box.innerHTML = html;
    const on = box.querySelector('.tab.on');
    if (on) box.scrollLeft = Math.max(0, on.offsetLeft - box.clientWidth / 2 + on.offsetWidth / 2);
  }
}
$('#tabs').addEventListener('click', (e) => {
  const b = e.target.closest('.tab');
  if (!b || b.dataset.tab === tab()) return;
  state.ui.tab = b.dataset.tab;
  save();
  state.ui.work = 0;
  if (typeOf(state.ui.tab) === 'asist') startCard();
  refresh();
  window.scrollTo(0, 0);
});

// Conteos del día (sin la fecha: ya está arriba) y el botón de vista previa del Excel, en cada pantalla.
export function renderSummary() {
  pintarConteos();
  $('#summary').insertAdjacentHTML('beforeend', '<button class="btn btn-previa" data-previa>Vista previa</button>');
}
$('#summary').addEventListener('click', (e) => {
  if (e.target.closest('[data-previa]')) import('./export.js').then((m) => m.openPreview('todo'));
});
function pintarConteos() {
  const cat = tab();
  const t = typeOf(cat);
  if (t === 'num') {
    const day = dayData(cat, false);
    const vals = lista().map((s) => day[s.row]).filter((v) => v != null);
    const avg = vals.length ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10 : null;
    $('#summary').innerHTML =
      `<span>${vals.length}/${lista().length} capturados</span>` +
      (avg != null ? `<span>Promedio: ${avg}</span>` : '');
    return;
  }
  if (t === 'disc') {
    const day = dayData(cat, false);
    const zeros = lista().filter((s) => day[s.row] === 0).length;
    $('#summary').innerHTML =
      `<span class="F">Con 0: ${zeros}</span><span>Los demás: 10</span>`;
    return;
  }
  if (t !== 'asist') {
    const day = dayData(cat, false);
    const vals = lista().map((s) => day[s.row]).filter((v) => v != null);
    $('#summary').innerHTML =
      `<span>${vals.length}/${lista().length} capturados</span>` +
      OPTIONS[t].map((o) => `<span class="${o.color}">${o.btn}: ${vals.filter((v) => v === o.v).length}</span>`).join('');
    return;
  }
  const rec = dayRecords(false);
  const counts = { A: 0, F: 0, R: 0, J: 0 };
  Object.values(rec).forEach((s) => counts[s]++);
  const marked = Object.keys(rec).length;
  $('#summary').innerHTML =
    `<span>${marked}/${lista().length} marcados</span>` +
    STATUSES.map((k) => `<span class="${k}">${k}: ${counts[k]}</span>`).join('');
}
// Cambió el grupo o la fecha: vuelve a empezar desde el primer pendiente.
export function changed() {
  state.ui.work = 0; // al cambiar de día o grupo se empieza en el primer trabajo
  save();
  startCard();
  refresh();
}
$('#group-select').addEventListener('change', (e) => {
  state.ui.group = Number(e.target.value);
  $('#search').value = '';
  changed();
});
// Fecha arriba: botón que abre el calendario (con los días que tuvieron actividad marcados).
const fechaLarga = (iso) => { const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d).toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' }); };
export function pintarFecha() {
  const b = $('#date-input');
  b.value = state.ui.date;
  b.textContent = fechaLarga(state.ui.date);
  b.classList.toggle('hoy', state.ui.date === todayISO());
}
export function irAFecha(iso) {
  state.ui.date = iso;
  pintarFecha();
  changed();
}
$('#date-input').addEventListener('click', () => import('./calendario.js').then((m) => m.abrirCalendario()));
export function shiftDay(n) {
  const [y, m, d] = state.ui.date.split('-').map(Number);
  const dt = new Date(y, m - 1, d + n);
  state.ui.date = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
  pintarFecha();
  changed();
}
$('#prev-day').addEventListener('click', () => shiftDay(-1));
$('#next-day').addEventListener('click', () => shiftDay(1));
$('#search').addEventListener('input', renderStudents);
