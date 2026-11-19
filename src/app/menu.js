import { openSettings } from './settings.js';
import { openAddStudent } from './students.js';
import { openSubject } from './subject.js';
import { confirmBulk } from './dialogs.js';
import { $, toast } from './util.js';
import { rolGuardado } from '../lib/nube.js';
import { DEMO } from '../lib/demo.js';
import { setVistaApp } from './state.js';
import { render } from './main-view.js';
import { alCambiarEstado, estado } from './sync.js';

// ---------- Menú y diálogos ----------
// Pantallas que se usan poco se cargan al abrirlas (la app abre más rápido).
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

menu.addEventListener('click', async (e) => {
  const action = e.target.dataset.action;
  if (!action || action === 'close') return;
  menu.close();
  if (action === 'add-student') openAddStudent();
  if (action === 'bajas') abrir(() => import('./bajas.js'), 'openBajas');
  if (action === 'grades') abrir(() => import('./grades.js'), 'openGrades');
  if (action === 'student') abrir(() => import('./student-report.js'), 'openStudent');
  if (action === 'group-report') abrir(() => import('./group-report.js'), 'openGroupReport');
  if (action === 'book') abrir(() => import('./book.js'), 'openBook');
  if (action === 'subject') openSubject();
  if (action === 'settings') openSettings();
  if (action === 'cuenta') openCuenta();
  if (action === 'escuela') { setVistaApp('escuela'); render(); }
  if (action === 'salir-demo') location.href = '/';
  if (action === 'semestral') abrir(() => import('./semestral.js'), 'openSemestral');
  if (action === 'new-file') {
    if (await confirmBulk('¿Cargar la lista nueva?', 'Elige el Excel con la lista actualizada (por ejemplo, la del nuevo parcial). Cada alumno conserva lo capturado y los parciales cerrados; los nuevos se agregan y los que ya no vengan quedan de baja. Antes de aplicar te digo quién llega y quién se va.', 'Elegir archivo')) $('#file-input').click();
  }
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

// Punto en el botón del menú: si lo capturado ya está guardado en la cuenta.
// Verde = guardado · azul = guardando · naranja = sin internet (se sube al volver) · rojo = hay un problema.
const PUNTO = {
  listo: ['ok', 'Todo guardado en tu cuenta'],
  subiendo: ['sube', 'Guardando en tu cuenta…'],
  pendiente: ['sube', 'Guardando en tu cuenta…'],
  'sin-conexion': ['espera', 'Sin internet: se guarda en tu cuenta al volver la conexión'],
  error: ['mal', 'No se pudo guardar en tu cuenta. Abre Mi cuenta para ver por qué'],
  'otra-cuenta': ['mal', 'Las listas de este celular son de otra cuenta'],
  vencida: ['mal', 'Tu acceso venció'],
};
function pintarPunto(e) {
  const [cls, txt] = PUNTO[e.fase] || ['', ''];
  [['#menu-btn', 'Menú'], ['#esc-cuenta', 'Mi cuenta']].forEach(([sel, nombre]) => {
    const b = $(sel);
    b.dataset.nube = cls;
    b.title = txt || nombre;
    b.setAttribute('aria-label', txt ? `${nombre}. ${txt}` : nombre);
  });
}
alCambiarEstado(pintarPunto);
pintarPunto(estado);
