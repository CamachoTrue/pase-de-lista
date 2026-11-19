import './lib/apariencia.js';
import { ASIST_ID, DEFAULT_CODES, normalizeGroup } from './lib/excel-core.js';
import { render } from './app/main-view.js';
import { fijarReferencia, fixGrading, state, todayISO } from './app/state.js';
import { DEMO } from './lib/demo.js';
import './app/util.js';
import './app/dialogs.js';
import './app/upload.js';
import './app/compact.js';
import './app/card.js';
import './app/students.js';
import './app/export.js';
import './app/subject.js';
import './app/menu.js';
import './app/settings.js';
import './app/sync.js';
import { alBloquear } from './app/acceso.js';

// ---------- Inicio ----------
if (state) {
  if (!state.ui) state.ui = { group: 0, date: todayISO() };
  if (state.ui.group >= state.groups.length) state.ui.group = 0;
  state.ui.date = todayISO(); // al abrir la app siempre empieza en hoy
  if (!state.ui.tab) state.ui.tab = ASIST_ID;
  state.codes = { ...DEFAULT_CODES, ...state.codes };
  // Versiones anteriores: porcentajes de todo el archivo y 5 secciones fijas. Cada grupo queda con su copia
  // y con sus secciones (las de siempre: Asistencia, Lecturas, Trabajos, Disciplina y Examen).
  state.groups.forEach((g) => {
    g.grading = fixGrading(g.grading || state.grading);
    normalizeGroup(g, g.grading);
    delete g.grading.lectMin; // el mínimo de "Sí" ahora es de cada sección
  });
  delete state.grading;
}
fijarReferencia();
// Sin cuenta: si se intenta capturar, el cambio se descarta y se vuelve a pintar lo guardado.
alBloquear(() => render());
// Demostración (otra pestaña): se nota en el título de la pestaña.
if (DEMO) document.title = `Demostración · ${document.title}`;
// Demostración nueva (pestaña recién abierta): sin cuadernillos de una demostración anterior.
if (DEMO && !state) try { indexedDB.deleteDatabase(`paseDeLista.demo.${DEMO}`); } catch { /* sin IndexedDB */ }
render();
// Cambios que llegaron de otro dispositivo.
window.addEventListener('nube-cambio', () => render());

// Enlace de invitación (?invitacion=TOKEN): abre la pantalla para crear la cuenta.
const invitacion = new URLSearchParams(location.search).get('invitacion');
if (!DEMO && invitacion && /^[0-9a-f-]{36}$/i.test(invitacion)) import('./app/cuenta.js').then((m) => m.openInvitacion(invitacion));

// En desarrollo (npm run dev) no hay service worker: la lista de archivos solo existe al construir.
if (import.meta.env.PROD && 'serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
