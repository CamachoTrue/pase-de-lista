// Escuela de ejemplo para el panel del director (/director?ejemplo=1): datos inventados, con la misma
// forma que regresa la base de datos, para enseñar el panel sin entrar ni tocar datos reales.
// Siempre salen iguales (generador con semilla) y las fechas se acomodan a hoy.
function azar(semilla) {
  let a = semilla;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const NOMBRES = ['ANA', 'LUIS', 'MARÍA', 'JOSÉ', 'SOFÍA', 'DIEGO', 'VALERIA', 'CARLOS', 'XIMENA', 'JUAN', 'CAMILA', 'MIGUEL', 'DANIELA', 'ÁNGEL',
  'FERNANDA', 'JESÚS', 'ANDREA', 'EMILIANO', 'PAOLA', 'RICARDO', 'REGINA', 'ALEJANDRO', 'LUCÍA', 'SEBASTIÁN', 'RENATA', 'ERICK', 'ITZEL', 'OMAR'];
const APELLIDOS = ['GARCÍA', 'LÓPEZ', 'MARTÍNEZ', 'VALENZUELA', 'RUIZ', 'SOTO', 'FÉLIX', 'ESPINOZA', 'MORALES', 'LEYVA', 'CASTRO', 'ROMERO',
  'ARMENTA', 'GÁMEZ', 'QUIJADA', 'BORBÓN', 'ACOSTA', 'MENDÍVIL', 'ZAZUETA', 'ORTIZ', 'ROJO', 'BACASEGUA', 'VEGA', 'IBARRA'];
const r1 = (n) => Math.round(n * 10) / 10;

export function escuelaEjemplo() {
  const rnd = azar(2026);
  const toma = (a) => a[Math.floor(rnd() * a.length)];
  const hoy = new Date();
  const diaMenos = (n) => new Date(hoy.getTime() - n * 864e5).toISOString().slice(0, 10);
  const maestros = [
    { id: 'm1', nombre: 'Prof. Jorge Méndez', correo: 'jorge.mendez@prepaejemplo.mx', rol: 'maestro', activo: true },
    { id: 'm2', nombre: 'Mtra. Sofía Navarro', correo: 'sofia.navarro@prepaejemplo.mx', rol: 'maestro', activo: true },
    { id: 'm3', nombre: 'Prof. Ricardo Soto', correo: 'ricardo.soto@prepaejemplo.mx', rol: 'maestro', activo: true },
    { id: 'm4', nombre: 'Mtra. Laura Gómez', correo: 'direccion@prepaejemplo.mx', rol: 'director', activo: true },
  ];
  // Grupo → materias con su maestro, escala y hace cuántos días capturó por última vez.
  const plan = {
    '1A': [['Matemáticas', 'm1', 10, 0], ['Literatura', 'm2', 10, 1], ['Inglés', 'm3', 100, 2]],
    '1B': [['Matemáticas', 'm1', 10, 0], ['Literatura', 'm2', 10, 3], ['Historia', 'm4', 10, 1]],
    '3A': [['Química', 'm1', 10, 1], ['Inglés', 'm3', 100, 11], ['Historia', 'm4', 10, 2]],
    '5B': [['Física', 'm1', 10, 0], ['Ética', 'm2', 10, 4]],
  };
  const alumnos = [];
  const clases = [];
  const resumenes = [];
  const y = hoy.getFullYear();
  let n = 0;
  for (const [grupo, materias] of Object.entries(plan)) {
    const lista = [];
    for (let i = 1; i <= 24; i++) {
      const id = `a${++n}`;
      alumnos.push({ id, nombre: `${toma(APELLIDOS)} ${toma(APELLIDOS)} ${toma(NOMBRES)}`, clave: `EJEMPLO-${grupo}-${String(i).padStart(3, '0')}` });
      // Qué tan bien va el alumno en general (casi todos bien, algunos batallan).
      lista.push({ id, nivel: rnd() < 0.12 ? 0.45 + rnd() * 0.15 : 0.65 + rnd() * 0.33, asiste: rnd() < 0.1 ? 0.7 + rnd() * 0.12 : 0.86 + rnd() * 0.14 });
    }
    materias.forEach(([materia, maestro, escala, atraso], k) => {
      const id = `c-${grupo}-${k}`;
      clases.push({ id, materia, maestro_id: maestro, grupo_id: grupo, grupos: { nombre: grupo } });
      const datos = { escala, alumnos: {}, semestre: `Ago–Dic ${y}`, parcialesCerrados: [1, 2] };
      for (const a of lista) {
        const nivel = Math.min(1, Math.max(0.3, a.nivel + (rnd() - 0.5) * 0.12));
        const asis = Math.min(1, Math.max(0.6, a.asiste + (rnd() - 0.5) * 0.06));
        const f = r1(nivel * escala);
        const att = r1(asis * 100);
        const faltas = Math.round((1 - asis) * 40);
        const pct = (f / escala) * 100;
        let e = 'verde';
        const r = [];
        if (pct < 60) { e = 'rojo'; r.push(`Calificación ${f} (reprobando)`); } else if (pct < 80) { e = 'amarillo'; r.push(`Calificación ${f}`); }
        if (att < 80) { e = 'rojo'; r.push(`Asistencia ${att} %`); } else if (att < 90) { if (e === 'verde') e = 'amarillo'; r.push(`Asistencia ${att} %`); }
        const p = [0, 1, 2].map(() => r1(Math.min(escala, Math.max(0, f + (rnd() - 0.5) * escala * 0.15))));
        datos.alumnos[a.id] = { f, a: att, F: faltas, e, r, p, s: r1((p[0] + p[1] + p[2]) / 3) };
      }
      resumenes.push({ clase_id: id, datos, ultima_captura: diaMenos(atraso), actualizado_en: new Date().toISOString() });
    });
  }
  return { clases, resumenes, alumnos, maestros };
}
