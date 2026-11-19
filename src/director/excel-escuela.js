// Excel de la escuela para el director: Resumen, Clases, Por maestro, En riesgo y una hoja por clase
// con sus alumnos. Sale de los mismos datos que se ven en el panel (los resúmenes de cada clase).
const r1 = (n) => (n == null ? null : Math.round(n * 10) / 10);
const avg = (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : null);
const sobre10 = (v, escala) => (v == null ? null : (v * 10) / (escala || 10));
const ESTADO = { rojo: 'En riesgo', amarillo: 'Observación', verde: 'Bien' };
const VERDE = 'FF1F7A4D';
const COLOR = { rojo: 'FFFBE5E5', amarillo: 'FFFBF0DC', verde: 'FFE3F4EA' };

function encabezado(ws, columnas) {
  ws.columns = columnas.map(([header, key, width]) => ({ header, key, width }));
  const h = ws.getRow(1);
  h.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE } };
  h.alignment = { vertical: 'middle', wrapText: true };
  h.height = 30;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columnas.length } };
}
const pintarEstado = (row, e, col) => {
  if (!COLOR[e]) return;
  row.getCell(col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR[e] } };
};

// Nombres de hoja: máximo 31 caracteres, sin \ / ? * [ ] : y sin repetir.
function nombreHoja(base, usados) {
  let n = String(base).replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Clase';
  for (let i = 2; usados.has(n.toUpperCase()); i++) n = `${n.slice(0, 27)} (${i})`;
  usados.add(n.toUpperCase());
  return n;
}

export async function excelEscuela(ExcelJS, d, nombre) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Clase al Día';
  const todos = d.clases.flatMap((c) => Object.entries(c.alumnos).map(([id, x]) => ({ id, ...x, c })));
  const alumno = (id) => d.alumnos[id] || {};
  const hoy = new Date().toLocaleDateString('es-MX', { day: 'numeric', month: 'long', year: 'numeric' });

  // Resumen
  const rs = wb.addWorksheet('Resumen');
  rs.getColumn(1).width = 34;
  rs.getColumn(2).width = 16;
  rs.addRow([nombre || 'Escuela']).font = { bold: true, size: 14 };
  rs.addRow([`Resumen al ${hoy}`]).font = { color: { argb: 'FF6B7570' } };
  rs.addRow([]);
  const unicos = new Set(todos.map((x) => x.id));
  const riesgo = new Set(todos.filter((x) => x.e === 'rojo').map((x) => x.id));
  const dias = (iso) => (iso ? Math.floor((Date.now() - Date.parse(`${iso}T12:00:00`)) / 864e5) : null);
  [
    ['Alumnos', unicos.size],
    ['Maestros', d.maestros.filter((m) => m.activo).length],
    ['Clases', d.clases.length],
    ['Asistencia promedio (%)', r1(avg(todos.map((x) => x.a).filter((v) => v != null)))],
    ['Promedio general (sobre 10)', r1(avg(todos.map((x) => sobre10(x.f, x.c.escala)).filter((v) => v != null)))],
    ['Alumnos en riesgo', riesgo.size],
    ['Clases sin captura en 7 días', d.clases.filter((c) => c.ultima == null || dias(c.ultima) > 7).length],
  ].forEach(([k, v]) => { const r = rs.addRow([k, v]); r.getCell(1).font = { bold: true }; });

  // Clases
  const cs = wb.addWorksheet('Clases');
  encabezado(cs, [['Grupo', 'g', 10], ['Materia', 'm', 22], ['Maestro', 'ma', 26], ['Alumnos', 'n', 10], ['Promedio', 'p', 11], ['Sobre', 'e', 8],
    ['Asistencia %', 'a', 13], ['En riesgo', 'r', 10], ['Observación', 'am', 12], ['Bien', 'v', 8], ['Última captura', 'u', 15]]);
  d.clases.forEach((c) => cs.addRow({ g: c.grupo, m: c.materia || '', ma: c.maestro, n: c.n, p: r1(c.promedio), e: c.escala, a: r1(c.asistencia),
    r: c.rojo, am: c.amarillo, v: c.verde, u: c.ultima || 'sin datos' }));

  // Por maestro
  const ms = wb.addWorksheet('Por maestro');
  encabezado(ms, [['Maestro', 'n', 26], ['Correo', 'c', 30], ['Clases', 'cl', 34], ['Alumnos', 'a', 10], ['En riesgo', 'r', 10], ['Última captura', 'u', 15], ['Acceso', 'ac', 10]]);
  d.maestros.forEach((m) => {
    const suyas = d.clases.filter((c) => c.maestro_id === m.id);
    ms.addRow({ n: m.nombre, c: m.correo, cl: suyas.map((c) => `${c.grupo}${c.materia ? ` ${c.materia}` : ''}`).join(', '),
      a: suyas.reduce((n, c) => n + c.n, 0), r: suyas.reduce((n, c) => n + c.rojo, 0),
      u: suyas.map((c) => c.ultima).filter(Boolean).sort().at(-1) || 'sin capturas', ac: m.activo ? 'Sí' : 'No' });
  });

  // En riesgo
  const rr = wb.addWorksheet('En riesgo');
  encabezado(rr, [['Alumno', 'n', 36], ['Clave', 'k', 18], ['Grupo', 'g', 9], ['Materia', 'm', 20], ['Maestro', 'ma', 24], ['Calif.', 'f', 9], ['Sobre', 'e', 8],
    ['Asistencia %', 'a', 13], ['Faltas', 'F', 8], ['Motivo', 'r', 44]]);
  todos.filter((x) => x.e === 'rojo').sort((a, b) => String(alumno(a.id).nombre).localeCompare(String(alumno(b.id).nombre))).forEach((x) => rr.addRow({
    n: alumno(x.id).nombre, k: alumno(x.id).clave, g: x.c.grupo, m: x.c.materia || '', ma: x.c.maestro, f: r1(x.f), e: x.c.escala, a: r1(x.a), F: x.F ?? 0, r: (x.r || []).join(' · '),
  }));

  // Una hoja por clase
  const usados = new Set(['RESUMEN', 'CLASES', 'POR MAESTRO', 'EN RIESGO']);
  for (const c of d.clases) {
    const ws = wb.addWorksheet(nombreHoja(`${c.grupo} ${c.materia || ''}`, usados));
    const np = Math.max(0, ...Object.values(c.alumnos).map((x) => (x.p || []).length));
    encabezado(ws, [['Alumno', 'n', 36], ['Clave', 'k', 18], [`Calif. (sobre ${c.escala})`, 'f', 12], ['Asistencia %', 'a', 13], ['Faltas', 'F', 8],
      ...Array.from({ length: np }, (_, i) => [`Parcial ${i + 1}`, `p${i}`, 10]), ...(np ? [['Semestral', 's', 11]] : []),
      ['Estado', 'e', 13], ['Motivo', 'r', 40]]);
    const filas = Object.entries(c.alumnos).map(([id, x]) => ({ id, ...x })).sort((a, b) => String(alumno(a.id).nombre).localeCompare(String(alumno(b.id).nombre)));
    for (const x of filas) {
      const row = ws.addRow({ n: alumno(x.id).nombre, k: alumno(x.id).clave, f: r1(x.f), a: r1(x.a), F: x.F ?? 0,
        ...Object.fromEntries((x.p || []).map((v, i) => [`p${i}`, r1(v)])), ...(np ? { s: r1(x.s) } : {}),
        e: ESTADO[x.e] || 'Sin datos', r: (x.r || []).join(' · ') });
      pintarEstado(row, x.e, ws.columns.findIndex((col) => col.key === 'e') + 1);
    }
  }
  return wb.xlsx.writeBuffer();
}

// En celular: compartir (WhatsApp, correo, Drive…). En computadora: descargar.
export async function entregar(buf, nombre) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], nombre, { type: blob.type });
  if (navigator.canShare && navigator.canShare({ files: [file] }) && /Android|iPhone|iPad/i.test(navigator.userAgent)) {
    await navigator.share({ files: [file], title: nombre }).catch(() => {});
    return;
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
