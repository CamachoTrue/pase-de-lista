import { doExport, openPreview } from './export.js';
import { openSettings } from './settings.js';
import { openAddStudent, openRemoveStudent } from './students.js';
import { openSubject } from './subject.js';
import { $, toast } from './util.js';

// ---------- Menú y diálogos ----------
// Pantallas que se usan poco se cargan al abrirlas (la app abre más rápido).
const abrir = (carga, fn) => carga().then((m) => m[fn]()).catch((e) => { console.error(e); toast('No se pudo abrir. Revisa tu conexión e inténtalo otra vez.'); });
export const openCuenta = () => abrir(() => import('./cuenta.js'), 'openCuenta');
export const menu = $('#menu');
$('#menu-btn').addEventListener('click', () => menu.showModal());
document.querySelectorAll('dialog').forEach((d) =>
  d.addEventListener('click', (e) => {
    if (e.target === d || e.target.dataset.action === 'close') d.close();
  })
);

menu.addEventListener('click', (e) => {
  const action = e.target.dataset.action;
  if (!action || action === 'close') return;
  menu.close();
  if (action === 'export') doExport(null);
  if (action === 'preview') openPreview();
  if (action === 'add-student') openAddStudent();
  if (action === 'remove-student') openRemoveStudent();
  if (action === 'grades') abrir(() => import('./grades.js'), 'openGrades');
  if (action === 'student') abrir(() => import('./student-report.js'), 'openStudent');
  if (action === 'group-report') abrir(() => import('./group-report.js'), 'openGroupReport');
  if (action === 'book') abrir(() => import('./book.js'), 'openBook');
  if (action === 'subject') openSubject();
  if (action === 'settings') openSettings();
  if (action === 'cuenta') openCuenta();
  if (action === 'semestral') abrir(() => import('./semestral.js'), 'openSemestral');
  if (action === 'new-file') {
    if (confirm('Se reemplazará la lista actual y su asistencia guardada. Exporta antes si la necesitas. ¿Continuar?')) {
      $('#file-input').click();
    }
  }
});

// Pantalla inicial (sin Excel): entrar con la cuenta para bajar los datos.
$('#upload-cuenta').addEventListener('click', openCuenta);
