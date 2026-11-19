// ---------- Acceso: para capturar hace falta una cuenta activa ----------
// Sin cuenta, o con la suscripción vencida, la app deja VER y EXPORTAR todo lo que hay en el celular
// (los datos nunca se borran), pero no capturar ni cambiar nada. La revisión está en save(): si no
// hay acceso, el cambio no se guarda y la pantalla vuelve a lo que estaba guardado.
import { DEMO } from '../lib/demo.js';
import { accesoVencido, haySesion } from '../lib/nube.js';
import { load, setGuardia, setState, state } from './state.js';
import { $ } from './util.js';

// Cuenta con la que se puede trabajar (o una demostración).
export const tieneAcceso = () => !!DEMO || (haySesion() && !accesoVencido());
// La lista de ejemplo se puede usar para conocer la app.
export const puedeCapturar = () => tieneAcceso() || !!(state && state.ejemplo);

export function avisarBloqueo() {
  const d = $('#bloqueo');
  const vencida = haySesion() && accesoVencido();
  $('#bloqueo-titulo').textContent = vencida ? 'Tu acceso venció' : 'Entra a tu cuenta para capturar';
  $('#bloqueo-texto').textContent = vencida
    ? 'Puedes ver y exportar todo lo que tienes, pero para seguir capturando hay que renovar. Si eres de una escuela, avisa a tu director.'
    : 'Para pasar lista y calificar necesitas tu cuenta. Lo que ya tienes en este celular no se pierde: al entrar se guarda en tu cuenta. Mientras, puedes verlo y exportarlo.';
  $('#bloqueo-cuenta').textContent = vencida ? 'Ver mi cuenta' : 'Entrar o crear mi cuenta';
  if (!d.open) d.showModal();
}

let vuelvaAPintar = () => {};
export const alBloquear = (fn) => { vuelvaAPintar = fn; };
setGuardia(() => puedeCapturar(), () => {
  setState(load()); // se descarta el cambio: queda lo que estaba guardado
  vuelvaAPintar();
  avisarBloqueo();
});

$('#bloqueo').addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  $('#bloqueo').close();
  if (b.id === 'bloqueo-cuenta') import('./menu.js').then((m) => m.openCuenta());
  if (b.id === 'bloqueo-exportar') import('./export.js').then((m) => m.doExport(null));
});
