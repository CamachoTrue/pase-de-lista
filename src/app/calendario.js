// ---------- Calendario de la fecha ----------
// Un mes a la vista. Los días en que se capturó algo en este grupo (cualquier sección) llevan un punto de
// color; así se ve de un vistazo qué días sí hubo clase registrada y cuáles quedaron vacíos.
import './calendario.css';
import { sectionsOf } from '../lib/excel-core.js';
import { group, state, todayISO } from './state.js';
import { $ } from './util.js';
import { irAFecha } from './main-view.js';

const d = () => $('#calendario');
const iso = (y, m, dia) => `${y}-${String(m + 1).padStart(2, '0')}-${String(dia).padStart(2, '0')}`;
const DIAS = ['L', 'M', 'M', 'J', 'V', 'S', 'D'];
let mes = null; // [año, mes 0-11]

// Días con algo capturado en el grupo, y si fue en la sección que se está viendo.
function actividad() {
  const g = group();
  const dias = new Map();
  for (const s of sectionsOf(g)) {
    for (const [fecha, dato] of Object.entries(g[s.id] || {})) {
      const hay = Array.isArray(dato) ? dato.some((w) => Object.keys(w.vals || {}).length) : Object.keys(dato || {}).length > 0;
      if (!hay) continue;
      const x = dias.get(fecha) || { secciones: 0, esta: false };
      x.secciones++;
      if (s.id === state.ui.tab) x.esta = true;
      dias.set(fecha, x);
    }
  }
  return dias;
}

function pintar() {
  const [y, m] = mes;
  const act = actividad();
  const hoy = todayISO();
  const primero = (new Date(y, m, 1).getDay() + 6) % 7; // lunes = 0
  const total = new Date(y, m + 1, 0).getDate();
  const conActividad = [...act.keys()].filter((f) => f.startsWith(iso(y, m, 1).slice(0, 8))).length;
  let celdas = '';
  for (let i = 0; i < primero; i++) celdas += '<span></span>';
  for (let dia = 1; dia <= total; dia++) {
    const f = iso(y, m, dia);
    const a = act.get(f);
    const cls = ['cal-dia', a ? 'con' : '', a && a.esta ? 'esta' : '', f === hoy ? 'hoy' : '', f === state.ui.date ? 'sel' : '', f > hoy ? 'futuro' : ''].filter(Boolean).join(' ');
    celdas += `<button class="${cls}" data-f="${f}" aria-label="${dia}${a ? ', con actividad' : ''}">${dia}</button>`;
  }
  const titulo = new Date(y, m, 1).toLocaleDateString('es-MX', { month: 'long', year: 'numeric' });
  d().innerHTML = `
    <div class="cal-head">
      <button class="icon-btn" data-mes="-1" aria-label="Mes anterior">‹</button>
      <div><h2>${titulo.charAt(0).toUpperCase() + titulo.slice(1)}</h2><p class="hint">${conActividad ? `${conActividad} día${conActividad === 1 ? '' : 's'} con actividad` : 'Sin actividad este mes'}</p></div>
      <button class="icon-btn" data-mes="1" aria-label="Mes siguiente">›</button>
    </div>
    <div class="cal-semana">${DIAS.map((x) => `<span>${x}</span>`).join('')}</div>
    <div class="cal-grid">${celdas}</div>
    <div class="cal-leyenda"><span><i class="cal-p con"></i> Con actividad</span><span><i class="cal-p"></i> Sin actividad</span><span><i class="cal-p hoy"></i> Hoy</span></div>
    <div class="row"><button class="btn" data-hoy>Ir a hoy</button><button class="btn" data-action="close">Cerrar</button></div>`;
}

export function abrirCalendario() {
  const [y, m] = state.ui.date.split('-').map(Number);
  mes = [y, m - 1];
  pintar();
  d().showModal();
}

d().addEventListener('click', (e) => {
  const dia = e.target.closest('[data-f]');
  if (dia) { d().close(); irAFecha(dia.dataset.f); return; }
  const cambio = e.target.closest('[data-mes]');
  if (cambio) {
    const nuevo = new Date(mes[0], mes[1] + Number(cambio.dataset.mes), 1);
    mes = [nuevo.getFullYear(), nuevo.getMonth()];
    pintar();
    return;
  }
  if (e.target.closest('[data-hoy]')) { d().close(); irAFecha(todayISO()); }
});
