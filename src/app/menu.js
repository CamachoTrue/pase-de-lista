import { $, toast } from './util.js';
import { setVistaApp, ui } from './state.js';
import { refresh, render } from './main-view.js';
import { alCambiarEstado, estado } from './sync.js';

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

// Punto en el ícono de cuenta: si lo capturado ya está guardado en la cuenta.
// Verde = guardado · azul = guardando · naranja = sin internet (se sube al volver) · rojo = hay un problema.
const PUNTO = {
  listo: ['ok', 'Todo guardado en tu cuenta'],
  subiendo: ['sube', 'Guardando en tu cuenta…'],
  pendiente: ['sube', 'Guardando en tu cuenta…'],
  'sin-conexion': ['espera', 'Sin internet: se guarda en tu cuenta al volver la conexión'],
  error: ['mal', 'No se pudo guardar en tu cuenta. Toca para ver por qué'],
  'otra-cuenta': ['mal', 'Las listas de este celular son de otra cuenta'],
  vencida: ['mal', 'Tu acceso venció'],
};
function pintarPunto(e) {
  const [cls, txt] = PUNTO[e.fase] || ['', 'Mi cuenta'];
  ['#menu-btn', '#esc-cuenta'].forEach((sel) => {
    const b = $(sel);
    b.dataset.nube = cls;
    b.title = txt;
    b.setAttribute('aria-label', `Mi cuenta. ${txt}`);
  });
}
alCambiarEstado(pintarPunto);
pintarPunto(estado);
