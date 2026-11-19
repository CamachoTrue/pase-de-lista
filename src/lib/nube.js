import { DEMO, rolDemo } from './demo.js';

// Conexión con Supabase. La librería (~40 KB) se carga solo cuando se usa la cuenta, no al abrir la app.
// La URL y la clave publicable son públicas por diseño: los permisos los cuida la base de datos (RLS).
export const SUPABASE_URL = 'https://nwgswikzkuuhhisdqbjq.supabase.co';
export const SUPABASE_KEY = 'sb_publishable_G5UgsI0Wc1xJgrMWYJW8Sg_pHiHdE5Y';
const SESION = 'cuenta.sesion';
const RECORDAR = 'cuenta.recordar';

const seguro = (fn) => { try { return fn(); } catch { return null; } };
// "Mantener la sesión iniciada": la sesión vive en localStorage (sigue al cerrar la app);
// si no, en sessionStorage (se borra al cerrar). Sin internet la sesión guardada sigue sirviendo.
export const recordar = () => seguro(() => localStorage.getItem(RECORDAR)) !== '0';
export const setRecordar = (v) => seguro(() => localStorage.setItem(RECORDAR, v ? '1' : '0'));
const almacen = {
  getItem: (k) => seguro(() => localStorage.getItem(k) ?? sessionStorage.getItem(k)),
  setItem: (k, v) => seguro(() => {
    (recordar() ? localStorage : sessionStorage).setItem(k, v);
    (recordar() ? sessionStorage : localStorage).removeItem(k);
  }),
  removeItem: (k) => seguro(() => { localStorage.removeItem(k); sessionStorage.removeItem(k); }),
};
// ¿Hay una sesión guardada? (sin cargar la librería)
export const haySesion = () => !DEMO && !!almacen.getItem(SESION);

// Rol de la cuenta (para abrir la pantalla correcta sin esperar a internet).
const ROL = 'cuenta.rol';
export const rolGuardado = () => (DEMO ? rolDemo() : haySesion() ? seguro(() => localStorage.getItem(ROL)) : null);
export const guardarRol = (r) => !DEMO && seguro(() => (r ? localStorage.setItem(ROL, r) : localStorage.removeItem(ROL)));
// Paquete de la cuenta: 'empresarial' = de una escuela (con director); 'individual' = independiente.
const PAQUETE = 'cuenta.paquete';
export const paqueteGuardado = () => (DEMO ? (DEMO === 'independiente' ? 'individual' : 'empresarial') : haySesion() ? seguro(() => localStorage.getItem(PAQUETE)) : null);
export const guardarPaquete = (p) => !DEMO && seguro(() => (p ? localStorage.setItem(PAQUETE, p) : localStorage.removeItem(PAQUETE)));

// Acceso de la cuenta (escuela activa y fecha de vencimiento), para saber sin internet si se puede capturar.
const ACCESO = 'cuenta.acceso';
export const guardarAcceso = (p) => !DEMO && seguro(() => (p ? localStorage.setItem(ACCESO, JSON.stringify({ activa: p.activa !== false, vence_el: p.vence_el || null, escuela: !!p.escuela_id, rol: p.rol }))
  : localStorage.removeItem(ACCESO)));
export function accesoVencido() {
  if (DEMO) return false;
  const a = seguro(() => JSON.parse(localStorage.getItem(ACCESO) || 'null'));
  if (!a || a.rol === 'superadmin' || !a.escuela) return false;
  return !a.activa || (!!a.vence_el && a.vence_el < new Date().toISOString().slice(0, 10));
}

// Identificador de este dispositivo (no es personal): la nube anota qué dispositivo cambió cada clase
// para que los demás sepan que hay algo nuevo que bajar, sin volver a bajar sus propios cambios.
const DISP = 'cuenta.dispositivo';
export function dispositivo() {
  let d = seguro(() => localStorage.getItem(DISP));
  if (!d) {
    d = (crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`).slice(0, 36);
    seguro(() => localStorage.setItem(DISP, d));
  }
  return d;
}

let cliente = null;
export async function nube() {
  if (DEMO) throw new Error('Demostración: sin conexión a la cuenta.');
  if (!cliente) {
    cliente = import('@supabase/supabase-js').then(({ createClient }) => createClient(SUPABASE_URL, SUPABASE_KEY, {
      auth: { storage: almacen, storageKey: SESION, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
      global: { headers: { 'x-dispositivo': dispositivo() } },
    })).catch((e) => { cliente = null; throw e; });
  }
  return cliente;
}

// Mensajes de error en español para lo más común.
export function errorLegible(e) {
  const m = String((e && (e.message || e.error_description)) || e || '');
  if (/Invalid login credentials/i.test(m)) return 'Correo o contraseña incorrectos.';
  if (/already registered|already been registered/i.test(m)) return 'Ese correo ya tiene cuenta. Escribe tu contraseña para entrar.';
  if (/Password should be at least/i.test(m)) return 'La contraseña debe tener al menos 8 caracteres.';
  if (/Email not confirmed/i.test(m)) return 'Falta confirmar el correo. Avisa al administrador.';
  if (/Failed to fetch|NetworkError|network/i.test(m)) return 'Sin conexión. Inténtalo cuando tengas internet.';
  if (/rate limit/i.test(m)) return 'Demasiados intentos. Espera un minuto.';
  return m.replace(/^.*?: /, '') || 'Algo salió mal. Inténtalo de nuevo.';
}
