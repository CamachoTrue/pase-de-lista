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
export const haySesion = () => !!almacen.getItem(SESION);

let cliente = null;
export async function nube() {
  if (!cliente) {
    cliente = import('@supabase/supabase-js').then(({ createClient }) => createClient(SUPABASE_URL, SUPABASE_KEY, {
      auth: { storage: almacen, storageKey: SESION, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
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
