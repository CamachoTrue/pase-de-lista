// Cuadernillos PDF guardados en el celular (IndexedDB, una entrada por grupo = nombre de la hoja).
// Aparte de la pantalla del cuadernillo para que la sincronización lo use sin cargar esa pantalla.
import { sufijoDemo } from '../lib/demo.js';

export const bookDB = (() => {
  let p;
  const open = () => (p = p || new Promise((res, rej) => {
    const r = indexedDB.open(`paseDeLista${sufijoDemo}`, 1);
    r.onupgradeneeded = () => r.result.createObjectStore('cuadernillos');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  }));
  const run = (mode, fn) => open().then((db) => new Promise((res, rej) => {
    const tx = db.transaction('cuadernillos', mode);
    const req = fn(tx.objectStore('cuadernillos'));
    tx.oncomplete = () => res(req && req.result);
    tx.onerror = () => rej(tx.error);
    tx.onabort = () => rej(tx.error);
  }));
  return {
    get: (k) => run('readonly', (s) => s.get(k)),
    put: (k, v) => run('readwrite', (s) => s.put(v, k)),
    del: (k) => run('readwrite', (s) => s.delete(k)),
  };
})();
