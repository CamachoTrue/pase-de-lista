// ---------- Apariencia: tema (automático, claro, oscuro) y color principal ----------
// Se guarda en el dispositivo y en la cuenta (ajustes), así sale igual en todos los dispositivos del maestro.
// El tema pone data-tema en <html>; el color reemplaza las variables --primary* (el verde de asistencia no cambia).
import { almacenApp, sufijoDemo } from './demo.js';

// Cada color con su versión para fondo claro y para fondo oscuro: [principal, al pasar el mouse, suave, texto encima].
export const COLORES = {
  verde: { nombre: 'Verde', claro: ['#1f7a4d', '#1a6640', '#e3f4ea', '#ffffff'], oscuro: ['#2f9e66', '#38b073', '#183a27', '#ffffff'] },
  azul: { nombre: 'Azul marino', claro: ['#1f4e8c', '#193f72', '#e4ecf7', '#ffffff'], oscuro: ['#6b9fe6', '#82b0ee', '#1b2a42', '#0b1424'] },
  vino: { nombre: 'Vino', claro: ['#8a2343', '#721c37', '#f6e3e9', '#ffffff'], oscuro: ['#e07a98', '#ea93ab', '#3a1c26', '#1e0a11'] },
  morado: { nombre: 'Morado', claro: ['#5b3a9e', '#4a2f83', '#ece6f7', '#ffffff'], oscuro: ['#a48be0', '#b39ce8', '#2a2140', '#140d24'] },
  terracota: { nombre: 'Terracota', claro: ['#b2532a', '#964521', '#f8e7de', '#ffffff'], oscuro: ['#ec9468', '#f2a782', '#3b2519', '#1f1008'] },
  grafito: { nombre: 'Grafito', claro: ['#33393f', '#24292e', '#e9eaec', '#ffffff'], oscuro: ['#b9c0c7', '#cdd3d9', '#262b30', '#111417'] },
};
const CLAVE = `app.apariencia${sufijoDemo}`;
const POR_DEFECTO = { tema: 'auto', color: 'verde' };

export function leerApariencia() {
  try { return { ...POR_DEFECTO, ...JSON.parse(almacenApp().getItem(CLAVE) || '{}') }; } catch { return { ...POR_DEFECTO }; }
}
export function guardarApariencia(a) {
  const nueva = { ...leerApariencia(), ...a };
  if (!COLORES[nueva.color]) nueva.color = 'verde';
  if (!['auto', 'claro', 'oscuro'].includes(nueva.tema)) nueva.tema = 'auto';
  try { almacenApp().setItem(CLAVE, JSON.stringify(nueva)); } catch { /* sin almacenamiento */ }
  aplicarApariencia();
  return nueva;
}

const sistemaOscuro = window.matchMedia('(prefers-color-scheme: dark)');
export function aplicarApariencia() {
  const a = leerApariencia();
  const oscuro = a.tema === 'oscuro' || (a.tema === 'auto' && sistemaOscuro.matches);
  const raiz = document.documentElement;
  raiz.dataset.tema = oscuro ? 'oscuro' : 'claro';
  const [p, hover, suave, texto] = COLORES[a.color][oscuro ? 'oscuro' : 'claro'];
  raiz.style.setProperty('--primary', p);
  raiz.style.setProperty('--primary-hover', hover);
  raiz.style.setProperty('--primary-soft', suave);
  raiz.style.setProperty('--primary-text', texto);
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = oscuro ? '#101412' : COLORES[a.color].claro[0];
}
sistemaOscuro.addEventListener('change', aplicarApariencia);
aplicarApariencia();
