import { refresh } from './main-view.js';
import { group, save } from './state.js';
import { $, toast } from './util.js';

// ---------- Materia del grupo ----------
// Para maestros que tienen el mismo grupo dos veces en el Excel (una hoja por materia).
export function openSubject() {
  const g = group();
  $('#subject-group').textContent = `Hoja del Excel: ${g.sheet} · ${g.students.length} alumnos`;
  $('#subject-input').value = g.subject || '';
  $('#subject-clear').hidden = !g.subject;
  $('#subject').showModal();
  setTimeout(() => $('#subject-input').focus(), 50);
}
export function setSubject(v) {
  const g = group();
  g.subject = v.replace(/\s+/g, ' ').trim().slice(0, 40);
  if (!g.subject) delete g.subject;
  save();
  $('#subject').close();
  refresh();
  toast(g.subject ? `${g.sheet} ahora dice “${g.subject}”` : `${g.sheet} quedó sin materia`);
}
$('#subject-form').addEventListener('submit', (e) => { e.preventDefault(); setSubject($('#subject-input').value); });
$('#subject-clear').addEventListener('click', () => setSubject(''));
