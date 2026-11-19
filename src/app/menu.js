import { $, toast } from './util.js';
import { setVistaApp, ui } from './state.js';
import { refresh, render } from './main-view.js';

// ---------- Mi cuenta, barra de pestañas y diálogos ----------
// La app tiene tres pestañas abajo: Pasar lista, Calificaciones y Grupo, y arriba «Mi cuenta». Las pantallas que se usan poco se cargan al abrirlas (la app abre más rápido).
export const abrir = (carga, fn, ...args) => carga().then((m) => m[fn](...args)).catch((e) => { console.error(e); toast('No se pudo abrir. Revisa tu conexión e inténtalo otra vez.'); });
export const openCuenta = () => abrir(() => import('./cuenta.js'), 'openCuenta');
// «Mi cuenta» (arriba a la derecha): cuenta, escuela del director y salir de la demostración.
$('#menu-btn').addEventListener('click', () => openCuenta());
document.querySelectorAll('dialog').forEach((d) =>
  d.addEventListener('click', (e) => {
    if (e.target === d || e.target.dataset.action === 'close') d.close();
  })
);

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
$('#upload-registro').addEventListener('click', () => abrir(() => import('./cuenta.js'), 'openRegistro'));
$('#upload-escuela').addEventListener('click', () => { setVistaApp('escuela'); render(); });

// Aviso de acceso / lista de ejemplo (arriba de la lista).
$('#aviso-acceso').addEventListener('click', (e) => {
  if (e.currentTarget.dataset.ir === 'excel') $('#file-input').click();
  else openCuenta();
});
