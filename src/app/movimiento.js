// ---------- Movimiento y respuesta al tocar ----------
// Capa de pulido: no cambia lo que hace la app, solo cómo se siente. Respeta «reducir movimiento».
import { $ } from './util.js';

const reducir = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

// 1. Al tocar: una onda que sale del dedo. Al capturar (A/F/R/J, Sí/No, 0/50/100), un clic de vibración
//    en Android (iPhone no deja vibrar desde una página web; ahí queda la onda).
const TOCABLES = '.btn, .big-choice, .choice, .cb, .tab, .per-tab, .gr-chip, .esc-tab, .dir-tile, .tabbar button, .btn-cuenta, .exp-op';
const CAPTURA = '.big-choice, .choice, .cb, .fill-rest';
document.addEventListener('pointerdown', (e) => {
  const el = e.target.closest(TOCABLES);
  if (!el || el.disabled) return;
  if (e.target.closest(CAPTURA) && navigator.vibrate) { try { navigator.vibrate(8); } catch { /* sin vibración */ } }
  if (reducir()) return;
  const r = el.getBoundingClientRect();
  const d = Math.max(r.width, r.height) * 2.2;
  const onda = document.createElement('span');
  onda.className = 'onda';
  onda.style.cssText = `width:${d}px;height:${d}px;left:${e.clientX - r.left - d / 2}px;top:${e.clientY - r.top - d / 2}px`;
  if (getComputedStyle(el).position === 'static') el.style.position = 'relative';
  el.classList.add('con-onda');
  el.appendChild(onda);
  onda.addEventListener('animationend', () => onda.remove());
}, { passive: true });

// 2. Barra de abajo: la pastilla verde se desliza hasta la pestaña elegida.
const barra = $('#tabbar');
let ultima = '';
function moverPastilla() {
  const on = barra.querySelector('button.on');
  if (!on || !on.offsetWidth) return;
  const pos = `${on.offsetLeft}|${on.offsetWidth}`;
  if (pos === ultima) return; // sin cambios: no tocar nada (evita volver a disparar al observador)
  ultima = pos;
  barra.style.setProperty('--px', `${on.offsetLeft}px`);
  barra.style.setProperty('--pw', `${on.offsetWidth}px`);
  if (!barra.classList.contains('con-pastilla')) barra.classList.add('con-pastilla');
}
// Solo cuando cambia cuál botón está marcado (no la barra misma ni las ondas).
new MutationObserver((cambios) => { if (cambios.some((c) => c.target !== barra)) moverPastilla(); })
  .observe(barra, { subtree: true, attributes: true, attributeFilter: ['class'] });
// Cuando la barra aparece o cambia de tamaño (al abrir la app, girar el celular), se acomoda la pastilla.
new ResizeObserver(() => { ultima = ''; moverPastilla(); }).observe(barra);
