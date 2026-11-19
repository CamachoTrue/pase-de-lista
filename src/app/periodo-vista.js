// ---------- De qué periodo es el resumen: todo el semestre o un parcial ----------
// Lo comparten el resumen del grupo y el del alumno (al pasar de uno al otro se queda el mismo periodo).
// Un parcial cerrado se ve con las calificaciones y los porcentajes con que se cerró.
import { enPeriodo, periodosListos } from '../lib/excel-core.js';
import { cierreActivo } from './cierres.js';
import { state } from './state.js';

let periodo = 'todo'; // 'todo' | número de parcial
const parciales = () => (periodosListos(state.periodos) ? state.periodos.parciales : []);
const fechaC = (iso) => new Date(`${iso}T12:00:00`).toLocaleDateString('es-MX', { day: 'numeric', month: 'short' });

// El grupo tal como se ve en el periodo elegido.
export function vistaPeriodo(g) {
  const p = parciales().find((x) => x.numero === periodo);
  if (!p) { periodo = 'todo'; return { g, grading: g.grading, cierre: null, etiqueta: parciales().length ? 'Todo el semestre' : '' }; }
  const cierre = cierreActivo(g, 'parcial', p.numero);
  return {
    g: enPeriodo(g, p.inicio, p.fin),
    grading: (cierre && cierre.criterios) || g.grading,
    cierre,
    etiqueta: `Parcial ${p.numero} (${fechaC(p.inicio)} – ${fechaC(p.fin)})${cierre ? ' · cerrado' : ''}`,
    sufijo: ` - parcial ${p.numero}`,
  };
}
// Calificación del alumno en el periodo (la del cierre si el parcial está cerrado).
export const califCerrada = (v, row) => (v.cierre && v.cierre.calificaciones[row] ? v.cierre.calificaciones[row].final : undefined);
// Todas las del cierre { fila: calificación } (o null si el parcial no está cerrado).
export const cerradasDe = (v) => (v.cierre ? Object.fromEntries(Object.entries(v.cierre.calificaciones).map(([r, c]) => [r, c.final])) : null);

// Botones para elegir el periodo (solo si hay parciales con fechas).
export function botonesPeriodo(g) {
  if (!parciales().length) return '';
  return [['todo', 'Todo el semestre'], ...parciales().map((p) => [p.numero, `Parcial ${p.numero}${cierreActivo(g, 'parcial', p.numero) ? ' 🔒' : ''}`])]
    .map(([k, t]) => `<button type="button" class="per-tab ${periodo === k ? 'on' : ''}" data-periodo="${k}">${t}</button>`).join('');
}
// Al tocar un botón: cambia el periodo y vuelve a pintar.
export function alElegirPeriodo(nav, pintar) {
  nav.addEventListener('click', (e) => {
    const b = e.target.closest('[data-periodo]');
    if (!b) return;
    periodo = b.dataset.periodo === 'todo' ? 'todo' : Number(b.dataset.periodo);
    pintar();
  });
}
