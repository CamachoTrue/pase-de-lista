// Pruebas de permisos (Row Level Security) con Postgres local (PGlite): dos escuelas ficticias,
// sus directores y maestros. Cada quien debe ver y cambiar solo lo que le toca. npm test
import { test, before } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { PGlite } from '@electric-sql/pglite';

const db = new PGlite();
const id = (n) => `00000000-0000-0000-0000-${String(n).padStart(12, '0')}`;
const U = {
  super: { id: id(1), email: 'super@prueba.mx' },
  dirA: { id: id(11), email: 'dir.a@prueba.mx' }, maeA1: { id: id(12), email: 'mae.a1@prueba.mx' }, maeA2: { id: id(13), email: 'mae.a2@prueba.mx' },
  dirB: { id: id(21), email: 'dir.b@prueba.mx' }, maeB: { id: id(22), email: 'mae.b@prueba.mx' },
  nuevo: { id: id(31), email: 'nuevo@prueba.mx' }, indiv: { id: id(41), email: 'indiv@prueba.mx' },
};
const ESC_A = id(100), ESC_B = id(200);

// Lo mínimo de Supabase para correr las migraciones fuera de Supabase.
const SUPABASE_STUB = `
  create role anon; create role authenticated;
  create schema auth;
  create table auth.users (id uuid primary key, email text);
  create function auth.jwt() returns jsonb language sql stable as
    $$ select coalesce(nullif(current_setting('request.jwt.claims', true), '')::jsonb, '{}'::jsonb) $$;
  create function auth.uid() returns uuid language sql stable as $$ select (auth.jwt() ->> 'sub')::uuid $$;
  grant usage on schema auth to anon, authenticated;
  grant usage on schema public to anon, authenticated;
  create schema storage;
  create table storage.buckets (id text primary key, name text, public boolean, file_size_limit bigint, allowed_mime_types text[]);
  create table storage.objects (id uuid primary key default gen_random_uuid(), bucket_id text, name text);
  create function storage.foldername(name text) returns text[] language sql immutable as
    $$ select (string_to_array(name, '/'))[1:array_length(string_to_array(name, '/'), 1) - 1] $$;
  alter table storage.objects enable row level security;
  grant usage on schema storage to authenticated;
  grant select, insert, update, delete on storage.objects to authenticated;
`;

// Corre una consulta como un usuario (o anon) y regresa filas; o el error si la rechaza.
async function as(user, sql, params = []) {
  await db.exec('reset role');
  await db.query("select set_config('request.jwt.claims', $1, false)", [user ? JSON.stringify({ sub: user.id, email: user.email }) : '']);
  await db.exec(`set role ${user ? 'authenticated' : 'anon'}`);
  try { return (await db.query(sql, params)).rows; } finally { await db.exec('reset role'); }
}
const fails = async (user, sql, params, re) => assert.rejects(() => as(user, sql, params), re);
const count = async (user, table) => (await as(user, `select count(*)::int n from ${table}`))[0].n;

before(async () => {
  await db.exec(SUPABASE_STUB);
  const dir = new URL('../supabase/migrations/', import.meta.url);
  for (const f of readdirSync(dir).filter((x) => x.endsWith('.sql')).sort()) await db.exec(readFileSync(new URL(f, dir), 'utf8'));
  // Datos de partida (como lo haría el superadmin desde el panel).
  for (const u of Object.values(U)) await db.query('insert into auth.users values ($1, $2)', [u.id, u.email]);
  await db.exec(`
    insert into escuelas (id, nombre, clave, cuentas_contratadas) values ('${ESC_A}', 'Prepa A', 'ESCA', 3), ('${ESC_B}', 'Prepa B', 'ESCB', 5);
    insert into usuarios (id, escuela_id, rol, nombre, correo) values
      ('${U.super.id}', null, 'superadmin', 'Super', '${U.super.email}'),
      ('${U.dirA.id}', '${ESC_A}', 'director', 'Dir A', '${U.dirA.email}'),
      ('${U.maeA1.id}', '${ESC_A}', 'maestro', 'Mae A1', '${U.maeA1.email}'),
      ('${U.maeA2.id}', '${ESC_A}', 'maestro', 'Mae A2', '${U.maeA2.email}'),
      ('${U.dirB.id}', '${ESC_B}', 'director', 'Dir B', '${U.dirB.email}'),
      ('${U.maeB.id}', '${ESC_B}', 'maestro', 'Mae B', '${U.maeB.email}');
  `);
  // El maestro A1 importa su Excel: grupo, alumnos, su clase, secciones y una captura.
  await as(U.maeA1, `insert into grupos (escuela_id, nombre) values ('${ESC_A}', 'V205')`);
  await as(U.maeA1, `insert into alumnos (escuela_id, clave, nombre) values ('${ESC_A}', 'ESCA-V205-001', 'ALUMNA UNO')`);
  await as(U.maeA1, `insert into inscripciones (escuela_id, grupo_id, alumno_id, numero_lista) select '${ESC_A}', g.id, a.id, '1' from grupos g, alumnos a`);
  await as(U.maeA1, `insert into clases (escuela_id, grupo_id, maestro_id, materia) select '${ESC_A}', id, '${U.maeA1.id}', 'Literatura' from grupos`);
  await as(U.maeA1, `insert into secciones (escuela_id, clase_id, nombre, plantilla, porcentaje) select '${ESC_A}', id, 'Asistencia', 'asist', 100 from clases`);
  await as(U.maeA1, `insert into capturas (escuela_id, clase_id, seccion_id, alumno_id, fecha, valor)
    select s.escuela_id, s.clase_id, s.id, a.id, '2026-11-18', 'A' from secciones s, alumnos a`);
  // La escuela B también tiene un alumno.
  await as(U.maeB, `insert into alumnos (escuela_id, clave, nombre) values ('${ESC_B}', 'ESCB-M401-001', 'ALUMNO B')`);
});

test('sin sesión (anon) no se puede leer nada', async () => {
  await fails(null, 'select * from alumnos', [], /permission denied/);
  await fails(null, 'select * from escuelas', [], /permission denied/);
});

test('una escuela no ve nada de otra', async () => {
  for (const t of ['alumnos', 'grupos', 'clases', 'capturas', 'secciones']) {
    const rows = await as(U.maeB, `select escuela_id from ${t}`);
    assert.ok(rows.every((r) => r.escuela_id === ESC_B), `maestro B vio ${t} de otra escuela`);
    const rowsD = await as(U.dirB, `select escuela_id from ${t}`);
    assert.ok(rowsD.every((r) => r.escuela_id === ESC_B), `director B vio ${t} de otra escuela`);
  }
  assert.equal(await count(U.maeB, 'escuelas'), 1);
  await fails(U.maeB, `insert into alumnos (escuela_id, clave, nombre) values ('${ESC_A}', 'X', 'INTRUSO')`, [], /row-level security/);
});

test('maestro: ve alumnos de su escuela, pero solo sus clases y capturas', async () => {
  assert.equal(await count(U.maeA2, 'alumnos'), 1);
  assert.equal(await count(U.maeA2, 'grupos'), 1);
  assert.equal(await count(U.maeA2, 'clases'), 0);
  assert.equal(await count(U.maeA2, 'capturas'), 0);
  assert.equal(await count(U.maeA1, 'clases'), 1);
  assert.equal(await count(U.maeA1, 'capturas'), 1);
  // No puede escribir en la clase de otro maestro.
  // Aunque conozca los ids exactos (no puede ni ver esas secciones).
  const [k] = (await db.query(`select s.clase_id, s.id seccion_id, a.id alumno_id from secciones s, alumnos a where a.escuela_id = '${ESC_A}'`)).rows;
  await fails(U.maeA2, `insert into capturas (escuela_id, clase_id, seccion_id, alumno_id, fecha, valor) values ('${ESC_A}', $1, $2, $3, '2026-11-19', 'F')`,
    [k.clase_id, k.seccion_id, k.alumno_id], /row-level security/);
  const upd = await as(U.maeA2, `update capturas set valor = 'F' returning id`);
  assert.equal(upd.length, 0);
});

test('director: ve todo lo de su escuela pero no captura en clases ajenas', async () => {
  assert.equal(await count(U.dirA, 'clases'), 1);
  assert.equal(await count(U.dirA, 'capturas'), 1);
  assert.equal(await count(U.dirA, 'usuarios'), 3);
  const upd = await as(U.dirA, `update capturas set valor = 'F' returning id`);
  assert.equal(upd.length, 0);
});

test('nadie cambia su rol, su escuela ni el precio; los pagos solo los ve quien dirige', async () => {
  await fails(U.maeA1, `update usuarios set rol = 'director' where id = '${U.maeA1.id}'`, [], /permission denied/);
  await fails(U.dirA, `update usuarios set escuela_id = '${ESC_B}' where id = '${U.dirA.id}'`, [], /permission denied/);
  const precio = await as(U.dirA, `update escuelas set precio_mensual = 1 returning id`);
  assert.equal(precio.length, 0);
  assert.equal((await as(U.maeA1, `update usuarios set nombre = 'Mtra. Ana' where id = '${U.maeA1.id}' returning nombre`))[0].nombre, 'Mtra. Ana');
  await db.query(`insert into pagos (escuela_id, periodo_inicio, periodo_fin, monto) values ('${ESC_A}', '2026-08-01', '2026-12-31', 1000)`);
  assert.equal(await count(U.dirA, 'pagos'), 1);
  assert.equal(await count(U.maeA1, 'pagos'), 0);
  assert.equal(await count(U.dirB, 'pagos'), 0);
});

test('no se puede mezclar un alumno de otra escuela en una captura', async () => {
  // La llave compuesta (alumno, escuela) lo impide aunque el maestro conozca el id del alumno.
  const [b] = await db.query(`select id from alumnos where escuela_id = '${ESC_B}'`).then((r) => r.rows);
  await assert.rejects(() => as(U.maeA1, `insert into capturas (escuela_id, clase_id, seccion_id, alumno_id, fecha, valor)
    select s.escuela_id, s.clase_id, s.id, $1, '2026-11-20', 'A' from secciones s`, [b.id]), /foreign key|violates/);
});

test('historial: queda quién cambió una calificación y nadie lo borra', async () => {
  await as(U.maeA1, `update capturas set valor = 'R'`);
  const h = await as(U.dirA, `select tabla, accion, usuario_id, antes->>'valor' antes, despues->>'valor' despues from historial where tabla = 'capturas'`);
  assert.deepEqual(h.map((r) => [r.accion, r.antes, r.despues, r.usuario_id]), [['update', 'A', 'R', U.maeA1.id]]);
  assert.equal((await as(U.maeA1, 'select * from capturas'))[0].actualizado_por, U.maeA1.id);
  await fails(U.maeA1, 'delete from historial', [], /permission denied/);
  await fails(U.dirA, 'delete from historial', [], /permission denied/);
  assert.equal(await count(U.dirB, 'historial'), 0);
});

test('invitaciones: respetan cuentas contratadas, el correo y un solo uso', async () => {
  // La escuela A contrató 3 y ya usa 3.
  await fails(U.dirA, `insert into invitaciones (escuela_id, correo) values ('${ESC_A}', '${U.nuevo.email}')`, [], /cuentas contratadas/);
  await db.query(`update escuelas set cuentas_contratadas = 4 where id = '${ESC_A}'`);
  await fails(U.maeA1, `insert into invitaciones (escuela_id, correo) values ('${ESC_A}', '${U.nuevo.email}')`, [], /row-level security/);
  await fails(U.dirA, `insert into invitaciones (escuela_id, correo, rol) values ('${ESC_A}', '${U.nuevo.email}', 'director')`, [], /row-level security/);
  const [inv] = await as(U.dirA, `insert into invitaciones (escuela_id, correo) values ('${ESC_A}', '${U.nuevo.email}') returning token`);
  await fails(U.maeB, `select aceptar_invitacion($1, 'X')`, [inv.token], /ya pertenece|otro correo/);
  const otro = { id: id(32), email: 'otro@prueba.mx' };
  await db.query('insert into auth.users values ($1, $2)', [otro.id, otro.email]);
  await fails(otro, `select aceptar_invitacion($1, 'X')`, [inv.token], /otro correo/);
  await as(U.nuevo, `select aceptar_invitacion($1, 'Mtro. Nuevo')`, [inv.token]);
  assert.equal((await as(U.nuevo, 'select rol, escuela_id from usuarios'))[0].escuela_id, ESC_A);
  await fails(U.nuevo, `select aceptar_invitacion($1, 'X')`, [inv.token], /ya pertenece|no es válida/);
});

test('el director desactiva a un maestro y este deja de ver todo', async () => {
  await fails(U.maeA1, `update usuarios set activo = false where id = '${U.maeA2.id}'`, [], /row-level security|director|0 rows/).catch(async () => {
    const r = await as(U.maeA1, `update usuarios set activo = false where id = '${U.maeA2.id}' returning id`);
    assert.equal(r.length, 0);
  });
  await as(U.dirA, `update usuarios set activo = false where id = '${U.maeA2.id}'`);
  assert.equal(await count(U.maeA2, 'alumnos'), 0);
  await fails(U.maeA2, `update usuarios set activo = true where id = '${U.maeA2.id}'`, [], /director/);
});

test('paquete individual: el maestro es su escuela y administra sus periodos', async () => {
  await as(U.indiv, `select registrar_individual('Mtra. Individual')`);
  const [me] = await as(U.indiv, 'select escuela_id, rol from usuarios');
  assert.equal(me.rol, 'maestro');
  const [esc] = await as(U.indiv, 'select paquete, clave from escuelas');
  assert.equal(esc.paquete, 'individual');
  await as(U.indiv, `insert into semestres (escuela_id, nombre, inicio, fin) values ($1, 'Ago–Dic 2026', '2026-08-17', '2026-12-18')`, [me.escuela_id]);
  // Un maestro de escuela empresarial no administra periodos.
  await fails(U.maeA1, `insert into semestres (escuela_id, nombre, inicio, fin) values ('${ESC_A}', 'S', '2026-08-17', '2026-12-18')`, [], /row-level security/);
  assert.equal(await count(U.dirA, 'semestres'), 0);
});

test('lo cerrado no se edita; el maestro no ve el precio de su escuela', async () => {
  await as(U.maeA1, `insert into cierres (escuela_id, clase_id, tipo, calificaciones, cerrado_por) select escuela_id, id, 'semestre', '{"a":9}', '${U.maeA1.id}' from clases`);
  await fails(U.maeA1, `update cierres set calificaciones = '{"a":10}'`, [], /permission denied/);
  await as(U.maeA1, `update cierres set reabierto_en = now(), reabierto_por = '${U.maeA1.id}'`);
  await fails(U.maeA1, 'delete from cierres', [], /permission denied/);
  await fails(U.maeA1, 'select precio_mensual from escuelas', [], /permission denied/);
  assert.equal((await as(U.maeA1, 'select nombre from escuelas'))[0].nombre, 'Prepa A');
  assert.equal((await as(U.maeA1, 'select * from datos_cobro()')).length, 0);
  assert.equal((await as(U.dirA, 'select * from datos_cobro()'))[0].cuentas_contratadas, 4);
  // Un cambio de nombre de alumno queda en el historial.
  await as(U.maeA1, `update alumnos set nombre = 'ALUMNA UNO CORREGIDA' where escuela_id = '${ESC_A}'`);
  assert.equal((await as(U.dirA, `select count(*)::int n from historial where tabla = 'alumnos'`))[0].n, 1);
});

test('enlace de invitación: muestra escuela y correo solo con el token; mi_perfil', async () => {
  const [inv] = await as(U.dirB, `insert into invitaciones (escuela_id, correo) values ('${ESC_B}', 'nueva@prueba.mx') returning token`);
  const [v] = await as(null, 'select * from ver_invitacion($1)', [inv.token]);
  assert.deepEqual([v.escuela, v.correo, v.rol, v.valida], ['Prepa B', 'nueva@prueba.mx', 'maestro', true]);
  assert.equal((await as(null, 'select * from ver_invitacion($1)', ['00000000-0000-0000-0000-000000000999'])).length, 0);
  await fails(null, 'select * from mi_perfil()', [], /permission denied/);
  const [p] = await as(U.maeB, 'select * from mi_perfil()');
  assert.deepEqual([p.escuela, p.rol, p.clave], ['Prepa B', 'maestro', 'ESCB']);
});

test('sincronizar: importar grupo reconoce alumnos y es repetible', async () => {
  const lista = [{ numero: '1', nombre: 'alumna uno corregida' }, { numero: '2', nombre: 'Bruno  Núñez' }];
  const [{ importar_grupo: r1 }] = await as(U.maeA1, `select importar_grupo('V205', 'Ética', $1::jsonb)`, [JSON.stringify(lista)]);
  const [existente] = (await db.query(`select id from alumnos where nombre = 'ALUMNA UNO CORREGIDA'`)).rows;
  assert.equal(r1.alumnos[0], existente.id);
  const [nuevo] = (await db.query('select clave, nombre from alumnos where id = $1', [r1.alumnos[1]])).rows;
  assert.deepEqual([nuevo.clave, nuevo.nombre], ['ESCA-V205-002', 'Bruno  Núñez']);
  const [{ importar_grupo: r2 }] = await as(U.maeA1, `select importar_grupo('V205', 'Ética', $1::jsonb)`, [JSON.stringify(lista)]);
  assert.deepEqual([r2.clase_id, r2.alumnos], [r1.clase_id, r1.alumnos]);
  assert.equal(await count(U.maeA1, 'clases'), 2); // Literatura y Ética
  // Otra escuela no puede importar a este grupo: crea el suyo.
  const [{ importar_grupo: rb }] = await as(U.maeB, `select importar_grupo('V205', '', $1::jsonb)`, [JSON.stringify(lista)]);
  assert.notEqual(rb.grupo_id, r1.grupo_id);
});

test('sincronizar: secciones, días y trabajos dejan el día tal como viene', async () => {
  const lista = [{ numero: '1', nombre: 'alumna uno corregida' }, { numero: '2', nombre: 'Bruno  Núñez' }];
  const [{ importar_grupo: r }] = await as(U.maeA1, `select importar_grupo('V205', 'Ética', $1::jsonb)`, [JSON.stringify(lista)]);
  const [a1, a2] = r.alumnos;
  const secs = [
    { local: 'records', nombre: 'Asistencia', plantilla: 'asist', porcentaje: 50, orden: 0 },
    { local: 'trab', nombre: 'Tareas', plantilla: 'trab', porcentaje: 50, orden: 1 },
  ];
  const [{ guardar_secciones: ids }] = await as(U.maeA1, 'select guardar_secciones($1, $2, 100, 10, $3::jsonb)', [r.clase_id, 'Ética', JSON.stringify(secs)]);
  const hist = async () => (await db.query(`select count(*)::int n from historial where tabla = 'capturas'`)).rows[0].n;
  const caps = async () => (await db.query('select alumno_id, valor from capturas where seccion_id = $1 order by valor', [ids.records])).rows;
  const dia = (valores) => JSON.stringify([{ seccion: ids.records, fecha: '2026-11-20', valores }]);
  const h0 = await hist();
  await as(U.maeA1, 'select guardar_dias($1::jsonb)', [dia({ [a1]: 'A', [a2]: 'F' })]);
  assert.equal((await caps()).length, 2);
  await as(U.maeA1, 'select guardar_dias($1::jsonb)', [dia({ [a1]: 'A', [a2]: 'F' })]);
  assert.equal(await hist(), h0, 'repetir lo mismo no debe ensuciar el historial');
  await as(U.maeA1, 'select guardar_dias($1::jsonb)', [dia({ [a1]: 'A', [a2]: 'R' })]);
  assert.equal(await hist(), h0 + 1);
  await as(U.maeA1, 'select guardar_dias($1::jsonb)', [dia({ [a1]: 'A' })]);
  assert.deepEqual((await caps()).map((x) => x.valor), ['A']);
  assert.equal(await hist(), h0 + 2);
  // Trabajos: dos el mismo día, luego se quita uno (y sus calificaciones).
  const tr = (trabajos) => JSON.stringify([{ seccion: ids.trab, fecha: '2026-11-20', trabajos }]);
  const [{ guardar_trabajos: [t] }] = await as(U.maeA1, 'select guardar_trabajos($1::jsonb)', [tr([{ titulo: 'Unit 1' }, { titulo: '' }])]);
  assert.equal(t.ids.length, 2);
  await as(U.maeA1, 'select guardar_dias($1::jsonb)', [JSON.stringify(t.ids.map((id, i) => ({ seccion: ids.trab, fecha: '2026-11-20', trabajo: id, valores: { [a1]: String(50 * (i + 1)) } })))]);
  assert.equal((await db.query('select count(*)::int n from capturas where seccion_id = $1', [ids.trab])).rows[0].n, 2);
  await as(U.maeA1, 'select guardar_trabajos($1::jsonb)', [tr([{ id: t.ids[0], titulo: 'Unit 1 (corregido)' }])]);
  assert.equal((await db.query('select count(*)::int n from capturas where seccion_id = $1', [ids.trab])).rows[0].n, 1);
  assert.equal((await db.query('select titulo from trabajos where id = $1', [t.ids[0]])).rows[0].titulo, 'Unit 1 (corregido)');
  // Renombrar e intercambiar secciones; quitar una borra sus capturas.
  const [{ guardar_secciones: ids2 }] = await as(U.maeA1, 'select guardar_secciones($1, $2, 50, 100, $3::jsonb)', [r.clase_id, 'Ética',
    JSON.stringify([{ id: ids.records, local: 'records', nombre: 'Asistencia', plantilla: 'asist', porcentaje: 100, orden: 0 }])]);
  assert.equal(ids2.records, ids.records);
  assert.equal((await db.query('select count(*)::int n from secciones where clase_id = $1', [r.clase_id])).rows[0].n, 1);
  // Otro maestro no puede escribir ni borrar en esa clase: no ve la sección, así que no cambia nada.
  const antes = await caps();
  await as(U.maeB, 'select guardar_dias($1::jsonb)', [dia({ [a1]: 'F' })]);
  await as(U.maeB, 'select guardar_dias($1::jsonb)', [dia({})]);
  assert.deepEqual(await caps(), antes);
});

test('archivos: cada maestro solo usa su propia carpeta', async () => {
  const mia = `${ESC_A}/${U.maeA1.id}/excel/lista.xlsx`;
  await as(U.maeA1, `insert into storage.objects (bucket_id, name) values ('archivos', $1)`, [mia]);
  await fails(U.maeA1, `insert into storage.objects (bucket_id, name) values ('archivos', $1)`, [`${ESC_A}/${U.dirA.id}/excel/x.xlsx`], /row-level security/);
  await fails(U.maeA1, `insert into storage.objects (bucket_id, name) values ('archivos', $1)`, [`${ESC_B}/${U.maeA1.id}/excel/x.xlsx`], /row-level security/);
  assert.equal((await as(U.maeA1, 'select name from storage.objects')).length, 1);
  assert.equal((await as(U.dirA, 'select name from storage.objects')).length, 0);
  assert.equal((await as(U.maeB, 'select name from storage.objects')).length, 0);
});

test('todas las tablas tienen RLS activo', async () => {
  const r = await db.query(`select relname from pg_class c join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public' and c.relkind = 'r' and not c.relrowsecurity`);
  assert.deepEqual(r.rows, []);
});
