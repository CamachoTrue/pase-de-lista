import { openExportar } from './export.js';
import { $, toast } from './util.js';
import { rolGuardado } from '../lib/nube.js';
import { DEMO } from '../lib/demo.js';
import { setVistaApp, ui } from './state.js';
import { refresh, render } from './main-view.js';

// ---------- Menú (⋮), barra de pestañas y diálogos ----------
// La app tiene tres pestañas abajo: Pasar lista, Calificaciones y Grupo. El ⋮ solo guarda lo de la
// cuenta y exportar. Las pantallas que se usan poco se cargan al abrirlas (la app abre más rápido).
export const abrir = (carga, fn, ...args) => carga().then((m) => m[fn](...args)).catch((e) => { console.error(e); toast('No se pudo abrir. Revisa tu conexión e inténtalo otra vez.'); });
export const openCuenta = () => abrir(() => import('./cuenta.js'), 'openCuenta');
export const menu = $('#menu');
$('#menu-btn').addEventListener('click', () => {
  $('#menu-escuela').hidden = rolGuardado() !== 'director';
  $('#menu-salir-demo').hidden = !DEMO;
  menu.showModal();
});
document.querySelectorAll('dialog').forEach((d) =>
  d.addEventListener('click', (e) => {
    if (e.target === d || e.target.dataset.action === 'close') d.close();
  })
);

menu.addEventListener('click', (e) => {
  const action = e.target.dataset.action;
  if (!action || action === 'close') return;
  menu.close();
  if (action === 'exportar') openExportar();
  if (action === 'cuenta') openCuenta();
  if (action === 'escuela') { setVistaApp('escuela'); render(); }
  if (action === 'salir-demo') location.href = '/';
});

// Barra de abajo: cambia de pestaña.
export function irA(vista) {
  ui.vista = vista;
  window.scrollTo(0, 0);
  refresh();
}
$('#tabbar').addEventListener('click', (e) => {
  const b = e.target.closest('[data-vista]');
  if (b) irA(b.dataset.vista);
});

// Pantalla inicial (sin Excel): entrar con la cuenta para bajar los datos.
$('#upload-cuenta').addEventListener('click', openCuenta);
$('#upload-escuela').addEventListener('click', () => { setVistaApp('escuela'); render(); });

// Aviso de acceso / lista de ejemplo (arriba de la lista).
$('#aviso-acceso').addEventListener('click', (e) => {
  if (e.currentTarget.dataset.ir === 'excel') $('#file-input').click();
  else openCuenta();
});
