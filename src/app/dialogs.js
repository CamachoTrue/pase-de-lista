import { ASIST_ID } from '../lib/excel-core.js';
import { optionOf } from './state.js';
import { $ } from './util.js';

// Confirmación propia de la app (los avisos del navegador pueden bloquearse en el celular).
export function confirmBulk(title, text, okText = 'Sí, ponerlo', danger = false) {
  return new Promise((resolve) => {
    const d = $('#confirm-bulk');
    $('#bulk-title').textContent = title;
    $('#bulk-text').textContent = text;
    $('#bulk-ok').textContent = okText;
    $('#bulk-ok').className = `btn ${danger ? 'danger' : 'primary'}`;
    d.returnValue = '';
    d.addEventListener('close', () => resolve(d.returnValue === 'ok'), { once: true });
    d.showModal();
  });
}

// Pide autorización antes de cambiar (o quitar) una marca que ya existía.
export function confirmChange(name, from, to, cat = ASIST_ID) {
  return new Promise((resolve) => {
    const d = $('#confirm-change');
    const pill = (sel, v) => {
      const el = $(sel);
      const o = optionOf(cat, v);
      el.className = `change-val ${o ? o.color : 'none'}`;
      el.textContent = o ? o.label : 'Sin marca';
    };
    $('#confirm-name').textContent = name;
    pill('#confirm-from', from);
    pill('#confirm-to', to);
    $('#confirm-ok').textContent = to != null ? 'Sí, cambiar' : 'Sí, quitar';
    d.returnValue = '';
    d.addEventListener('close', () => resolve(d.returnValue === 'ok'), { once: true });
    d.showModal();
  });
}
