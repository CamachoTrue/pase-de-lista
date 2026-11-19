// ---------- Cierres: calificaciones definitivas de un parcial o del semestre ----------
// Al cerrar se guardan las calificaciones tal como están; ya no cambian aunque se muevan datos.
// Para corregir se reabre (queda registrado cuándo). Se sincronizan con la tabla "cierres".
import { save } from './state.js';

// Cierre vigente (no reabierto) de un parcial (n) o del semestre (n = null).
export const cierreActivo = (g, tipo, n = null) =>
  (g.cierres || []).filter((c) => c.tipo === tipo && (c.parcial ?? null) === n && !c.reabierto_en).at(-1) || null;

// criterios: copia de los porcentajes con que se cerró (para exportar ese parcial igual después,
// aunque el siguiente parcial cambie de secciones o porcentajes).
export function cerrar(g, tipo, n, calificaciones, criterios = null) {
  g.cierres = g.cierres || [];
  const c = { tipo, parcial: n, calificaciones, cerrado_en: new Date().toISOString(), reabierto_en: null,
    ...(criterios ? { criterios: JSON.parse(JSON.stringify(criterios)) } : {}) };
  g.cierres.push(c);
  save();
  return c;
}

export function reabrir(c) {
  c.reabierto_en = new Date().toISOString();
  save();
}

export const fechaCorta = (iso) => new Date(iso).toLocaleDateString('es-MX', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' });
