import { READ, STATUSES } from '../lib/excel-core.js';
import { save, state } from './state.js';
import { $, toast } from './util.js';

export function openSettings() {
  const f = $('#settings-form');
  [...STATUSES, ...READ].forEach((k) => (f.elements[k].value = state.codes[k]));
  $('#settings').showModal();
}
$('#settings').addEventListener('close', () => {
  if ($('#settings').returnValue !== 'save') return;
  const f = $('#settings-form');
  const vals = (keys) => keys.map((k) => f.elements[k].value.trim());
  if ([...vals(STATUSES), ...vals(READ)].some((v) => v === '')) {
    toast('Ninguna opción puede quedar vacía');
    return;
  }
  // Asistencia puede repetir valores (p. ej. A = R = J = 1); Sí y No deben ser distintos.
  const [si, no] = vals(READ);
  if (si.toUpperCase() === no.toUpperCase()) {
    toast('Sí y No necesitan valores distintos');
    return;
  }
  [...STATUSES, ...READ].forEach((k) => (state.codes[k] = f.elements[k].value.trim()));
  save();
  toast('Valores guardados');
});
