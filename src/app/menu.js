import { openBook } from './book.js';
import { doExport, openPreview } from './export.js';
import { openGrades } from './grades.js';
import { openGroupReport } from './group-report.js';
import { openSettings } from './settings.js';
import { openStudent } from './student-report.js';
import { openAddStudent, openRemoveStudent } from './students.js';
import { openSubject } from './subject.js';
import { $ } from './util.js';

// ---------- Menú y diálogos ----------
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
  if (action === 'grades') openGrades();
  if (action === 'student') openStudent();
  if (action === 'group-report') openGroupReport();
  if (action === 'book') openBook();
  if (action === 'subject') openSubject();
  if (action === 'settings') openSettings();
  if (action === 'new-file') {
    if (confirm('Se reemplazará la lista actual y su asistencia guardada. Exporta antes si la necesitas. ¿Continuar?')) {
      $('#file-input').click();
    }
  }
});
