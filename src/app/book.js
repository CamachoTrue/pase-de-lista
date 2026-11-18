import { confirmBulk } from './dialogs.js';
import { gname, group } from './state.js';
import { $, toast } from './util.js';

// ---------- Cuadernillo (PDF por grupo) ----------
// El PDF es grande para localStorage: se guarda en IndexedDB, una entrada por grupo (nombre de la hoja),
// así sigue ahí aunque se vuelva a subir el Excel. Se muestra con PDF.js para que se vea igual en Android y iPhone.
// La sincronización se entera de que cambió un cuadernillo (sin importar sync.js aquí, para no crear un ciclo).
const avisarCuadernillo = () => window.dispatchEvent(new Event('cuadernillo-cambio'));
export const bookDB = (() => {
  let p;
  const open = () => (p = p || new Promise((res, rej) => {
    const r = indexedDB.open('paseDeLista', 1);
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
export let pdfjsReady = null;
export const loadPdfJs = () => (pdfjsReady = pdfjsReady || new Promise((res, rej) => {
  const sc = document.createElement('script');
  sc.src = 'vendor/pdfjs/pdf.min.js';
  sc.onload = () => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js'; res(window.pdfjsLib); };
  sc.onerror = () => { pdfjsReady = null; rej(new Error('No se pudo cargar el visor de PDF')); };
  document.head.appendChild(sc);
}));

export const ZOOMS = [1, 1.25, 1.5, 2, 2.5, 3];
export const book = { key: null, doc: null, rec: null, zoom: 0, page: 1, observer: null, saveTimer: null, gen: 0 };
export const fmtSize = (n) => (n > 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);

export function bookShow(mode) {
  $('#book-empty').hidden = mode !== 'empty';
  $('#book-bar').hidden = mode !== 'doc';
  $('#book-pages').hidden = mode !== 'doc';
  $('#book-actions').hidden = mode !== 'doc';
  $('#book-loading').hidden = mode !== 'loading';
}
export function closeBookDoc() {
  book.gen++;
  if (book.observer) book.observer.disconnect();
  book.observer = null;
  if (book.doc) book.doc.destroy();
  book.doc = null;
  $('#book-pages').innerHTML = '';
}
export async function openBook() {
  const g = group();
  book.key = g.sheet;
  closeBookDoc();
  $('#book-sub').textContent = gname(g);
  $('#book').showModal();
  bookShow('loading');
  let rec;
  try { rec = await bookDB.get(book.key); } catch (err) { console.error(err); }
  if (!rec) {
    $('#book-empty-text').textContent = `Aún no hay cuadernillo para ${gname(g)}. Sube el PDF una sola vez y aquí lo verás siempre, también sin internet.`;
    bookShow('empty');
    return;
  }
  showBookDoc(rec);
}
export async function showBookDoc(rec) {
  const gen = ++book.gen;
  bookShow('loading');
  book.rec = rec;
  try {
    const pdfjsLib = await loadPdfJs();
    // PDF.js se queda con el buffer que recibe: se le pasa una copia.
    const doc = await pdfjsLib.getDocument({ data: new Uint8Array(rec.data.slice(0)) }).promise;
    if (gen !== book.gen) { doc.destroy(); return; }
    book.doc = doc;
  } catch (err) {
    console.error(err);
    toast('No se pudo abrir el PDF');
    $('#book-empty-text').textContent = 'El PDF guardado no se pudo abrir. Sube otro.';
    bookShow('empty');
    return;
  }
  $('#book-sub').textContent = `${gname(group())} · ${rec.name} · ${fmtSize(rec.size)}`;
  $('#book-total').textContent = book.doc.numPages;
  bookShow('doc');
  await layoutBook(Math.min(rec.page || 1, book.doc.numPages));
}
// Coloca un espacio por página (con su proporción) y dibuja solo las que están cerca de la vista.
export async function layoutBook(goTo) {
  const gen = book.gen;
  const box = $('#book-pages');
  if (book.observer) book.observer.disconnect();
  const first = await book.doc.getPage(1);
  if (gen !== book.gen) return;
  const vp = first.getViewport({ scale: 1 });
  const width = Math.floor((box.clientWidth - 16) * ZOOMS[book.zoom]);
  let html = '';
  for (let n = 1; n <= book.doc.numPages; n++) {
    html += `<div class="book-sheet" data-n="${n}" style="width:${width}px;height:${Math.round((width * vp.height) / vp.width)}px"><span>${n}</span></div>`;
  }
  box.innerHTML = html;
  $('#book-zoom-val').textContent = `${Math.round(ZOOMS[book.zoom] * 100)} %`;
  $('#book-fpage').textContent = `${book.page} / ${book.doc.numPages}`;
  book.observer = new IntersectionObserver((entries) => entries.forEach((e) => {
    if (e.isIntersecting) drawPage(e.target, gen);
    else { const c = e.target.querySelector('canvas'); if (c) { c.width = 0; c.remove(); } } // libera memoria
  }), { root: box, rootMargin: '800px 0px' });
  box.querySelectorAll('.book-sheet').forEach((el) => book.observer.observe(el));
  goPage(goTo);
}
export async function drawPage(el, gen) {
  if (el.querySelector('canvas') || el.dataset.busy) return;
  el.dataset.busy = '1';
  try {
    const page = await book.doc.getPage(Number(el.dataset.n));
    if (gen !== book.gen) return;
    const css = el.clientWidth;
    const vp1 = page.getViewport({ scale: 1 });
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const vp = page.getViewport({ scale: (css / vp1.width) * dpr });
    const c = document.createElement('canvas');
    c.width = Math.floor(vp.width);
    c.height = Math.floor(vp.height);
    el.style.height = `${Math.round(vp.height / dpr)}px`; // por si la página tiene otro tamaño
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    if (gen !== book.gen || el.querySelector('canvas')) return;
    el.appendChild(c);
  } catch (err) {
    if (gen === book.gen) console.error(err);
  } finally {
    delete el.dataset.busy;
  }
}
export function goPage(n) {
  const box = $('#book-pages');
  const el = box.querySelector(`.book-sheet[data-n="${n}"]`);
  if (!el) return;
  box.scrollTop = el.offsetTop - box.offsetTop - 8;
  setBookPage(n);
}
export function setBookPage(n) {
  book.page = n;
  if (document.activeElement !== $('#book-page')) $('#book-page').value = n;
  if (book.doc) $('#book-fpage').textContent = `${n} / ${book.doc.numPages}`;
  // Recuerda la página para abrir ahí la próxima vez.
  clearTimeout(book.saveTimer);
  const rec = book.rec;
  const key = book.key;
  book.saveTimer = setTimeout(() => {
    if (!rec || rec.page === n) return;
    rec.page = n;
    bookDB.put(key, rec).catch(() => {});
  }, 800);
}
$('#book-pages').addEventListener('scroll', () => {
  const box = $('#book-pages');
  const mid = box.scrollTop + box.clientHeight / 3 + box.offsetTop;
  let n = 1;
  for (const el of box.children) { if (el.offsetTop <= mid) n = Number(el.dataset.n); else break; }
  if (n !== book.page) setBookPage(n);
}, { passive: true });
$('#book-page').addEventListener('change', (e) => {
  const n = Math.round(Number(e.target.value));
  if (!book.doc || !(n >= 1 && n <= book.doc.numPages)) { e.target.value = book.page; toast(`Escribe una página del 1 al ${book.doc ? book.doc.numPages : 1}`); return; }
  e.target.blur();
  goPage(n);
});
$('#book-page').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.dispatchEvent(new Event('change')); });
export const zoomBook = (d) => {
  const z = Math.max(0, Math.min(ZOOMS.length - 1, book.zoom + d));
  if (z === book.zoom || !book.doc) return;
  book.zoom = z;
  book.gen++;
  layoutBook(book.page);
};
$('#book-zoom-in').addEventListener('click', () => zoomBook(1));
$('#book-zoom-out').addEventListener('click', () => zoomBook(-1));
// Pantalla completa: solo el PDF, con controles flotantes mínimos.
// En Android usa la pantalla completa del sistema; en iPhone (no la permite) ocupa toda la ventana de la app.
export const bookDlg = $('#book');
export const isReading = () => bookDlg.classList.contains('reading');
export async function setReading(on) {
  if (on === isReading() || !book.doc) return;
  const page = book.page;
  bookDlg.classList.toggle('reading', on);
  try {
    if (on && bookDlg.requestFullscreen && !document.fullscreenElement) await bookDlg.requestFullscreen({ navigationUI: 'hide' });
    if (!on && document.fullscreenElement) await document.exitFullscreen();
  } catch (err) { /* sin pantalla completa del sistema: basta con ocultar lo demás */ }
  await new Promise((r) => requestAnimationFrame(() => r()));
  book.gen++;
  layoutBook(page); // el ancho cambió: se vuelven a acomodar las páginas en la misma página
}
$('#book-full').addEventListener('click', () => setReading(true));
$('#book-exit').addEventListener('click', () => setReading(false));
$('#book-fzoom-in').addEventListener('click', () => zoomBook(1));
$('#book-fzoom-out').addEventListener('click', () => zoomBook(-1));
// Botón "atrás"/Esc del sistema: primero sale de pantalla completa, no cierra el cuadernillo.
document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && isReading()) setReading(false); });
bookDlg.addEventListener('cancel', (e) => { if (isReading()) { e.preventDefault(); setReading(false); } });
$('#book-upload').addEventListener('click', () => $('#book-file').click());
$('#book-change').addEventListener('click', () => $('#book-file').click());
$('#book-file').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  e.target.value = '';
  if (!file) return;
  if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') { toast('Elige un archivo PDF'); return; }
  const rec = { name: file.name, size: file.size, data: await file.arrayBuffer(), page: 1, added: Date.now() };
  if (String.fromCharCode(...new Uint8Array(rec.data.slice(0, 5))) !== '%PDF-') { toast('Ese archivo no es un PDF válido'); return; }
  try {
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
    await bookDB.put(book.key, rec);
    avisarCuadernillo();
  } catch (err) {
    console.error(err);
    toast('No hay espacio suficiente en el dispositivo para guardar el PDF');
    return;
  }
  closeBookDoc();
  book.zoom = 0;
  toast(`Cuadernillo guardado para ${gname(group())}`);
  showBookDoc(rec);
});
$('#book-remove').addEventListener('click', async () => {
  const key = book.key;
  const name = book.rec ? book.rec.name : 'el cuadernillo';
  if (!(await confirmBulk(`¿Quitar el cuadernillo de ${gname(group())}?`, `Se borra “${name}” de este dispositivo. Puedes volver a subirlo cuando quieras.`, 'Sí, quitar', true))) return;
  await bookDB.del(key).catch(() => {});
  avisarCuadernillo();
  closeBookDoc();
  book.rec = null;
  $('#book-sub').textContent = gname(group());
  $('#book-empty-text').textContent = `Aún no hay cuadernillo para ${gname(group())}. Sube el PDF una sola vez y aquí lo verás siempre, también sin internet.`;
  bookShow('empty');
  toast('Cuadernillo quitado');
});
$('#book').addEventListener('close', () => { $('#book').classList.remove('reading'); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); clearTimeout(book.saveTimer); if (book.rec && book.rec.page !== book.page && book.doc) { book.rec.page = book.page; bookDB.put(book.key, book.rec).catch(() => {}); } closeBookDoc(); });
