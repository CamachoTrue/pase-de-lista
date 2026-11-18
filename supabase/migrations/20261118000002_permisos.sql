-- Permisos (Row Level Security), historial de cambios e invitaciones.
--   superadmin: todo · director: su escuela (lee todo, administra maestros y periodos, no captura)
--   maestro: lee su escuela (alumnos, grupos, periodos) y solo escribe en sus clases
--   paquete individual: el maestro es la escuela entera, así que además administra sus periodos.
-- Las reglas están en la base de datos: aunque alguien manipule la app, no puede leer otra escuela.

-- ---------- Funciones de apoyo (esquema privado, no expuesto por la API) ----------
create schema if not exists privado;
revoke all on schema privado from public;
grant usage on schema privado to authenticated;

-- security definer: leen usuarios sin pasar por RLS (evita recursión); search_path fijo por seguridad.
create or replace function privado.yo() returns public.usuarios
language sql stable security definer set search_path = '' as $$
  select u from public.usuarios u where u.id = auth.uid() and u.activo
$$;

create or replace function privado.mi_escuela() returns uuid
language sql stable security definer set search_path = '' as $$
  select u.escuela_id from public.usuarios u where u.id = auth.uid() and u.activo
$$;

create or replace function privado.es_superadmin() returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.usuarios u where u.id = auth.uid() and u.activo and u.rol = 'superadmin')
$$;

-- Puede ver todo lo de una escuela: su director, o el superadmin.
create or replace function privado.dirige(e uuid) returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (
    select 1 from public.usuarios u
    where u.id = auth.uid() and u.activo
      and (u.rol = 'superadmin' or (u.rol = 'director' and u.escuela_id = e))
  )
$$;

-- Puede administrar periodos de una escuela: quien la dirige, o el maestro del paquete individual.
create or replace function privado.administra(e uuid) returns boolean
language sql stable security definer set search_path = '' as $$
  select privado.dirige(e) or exists (
    select 1 from public.usuarios u join public.escuelas s on s.id = u.escuela_id
    where u.id = auth.uid() and u.activo and u.escuela_id = e and s.paquete = 'individual'
  )
$$;

-- La clase es del maestro que la da.
create or replace function privado.es_mi_clase(c uuid) returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (
    select 1 from public.clases k join public.usuarios u on u.id = k.maestro_id
    where k.id = c and k.maestro_id = auth.uid() and u.activo
  )
$$;

create or replace function privado.clase_de_seccion(s uuid) returns uuid
language sql stable security definer set search_path = '' as $$
  select clase_id from public.secciones where id = s
$$;

grant execute on all functions in schema privado to authenticated;

-- Datos de cobro de la escuela para quien la dirige (o el maestro individual, que es su escuela).
create or replace function public.datos_cobro()
returns table (precio_mensual numeric, cuentas_contratadas int, cuentas_usadas int, vence_el date)
language sql stable security definer set search_path = '' as $$
  select s.precio_mensual, s.cuentas_contratadas,
         (select count(*) from public.usuarios u where u.escuela_id = s.id and u.activo and u.rol <> 'superadmin')::int,
         s.vence_el
  from public.escuelas s
  where s.id = privado.mi_escuela() and privado.administra(s.id)
$$;
revoke execute on function public.datos_cobro() from public, anon;
grant execute on function public.datos_cobro() to authenticated;

-- ---------- Activar RLS en todas las tablas ----------
alter table public.escuelas enable row level security;
alter table public.usuarios enable row level security;
alter table public.invitaciones enable row level security;
alter table public.pagos enable row level security;
alter table public.semestres enable row level security;
alter table public.parciales enable row level security;
alter table public.alumnos enable row level security;
alter table public.grupos enable row level security;
alter table public.inscripciones enable row level security;
alter table public.clases enable row level security;
alter table public.secciones enable row level security;
alter table public.trabajos enable row level security;
alter table public.capturas enable row level security;
alter table public.cierres enable row level security;
alter table public.cuadernillos enable row level security;
alter table public.historial enable row level security;

-- Nadie sin sesión (anon) toca nada.
revoke all on all tables in schema public from anon;
-- Tablas futuras: tampoco para anon (cada migración nueva debe activar RLS en sus tablas).
alter default privileges in schema public revoke all on tables from anon;
grant select, insert, update, delete on all tables in schema public to authenticated;

-- ---------- Escuelas: las ve quien pertenece; solo el superadmin las crea o cambia (precios, cuentas) ----------
create policy escuelas_ver on public.escuelas for select to authenticated
  using (id = privado.mi_escuela() or privado.es_superadmin());
-- Columnas de cobro (precio, cuentas, vencimiento): el maestro no las lee; la app las pide con
-- privado.datos_cobro() solo para quien dirige. Las demás columnas sí.
revoke select on public.escuelas from authenticated;
grant select (id, nombre, clave, paquete, activa, creado_en) on public.escuelas to authenticated;
create policy escuelas_superadmin on public.escuelas for all to authenticated
  using (privado.es_superadmin()) with check (privado.es_superadmin());

-- ---------- Usuarios ----------
create policy usuarios_ver on public.usuarios for select to authenticated
  using (id = auth.uid() or privado.dirige(escuela_id) or privado.es_superadmin());
-- El director activa o desactiva maestros de su escuela; el superadmin todo. Altas: por invitación.
create policy usuarios_director on public.usuarios for update to authenticated
  using (privado.dirige(escuela_id) and rol <> 'superadmin')
  with check (privado.dirige(escuela_id) and rol <> 'superadmin');
create policy usuarios_superadmin on public.usuarios for all to authenticated
  using (privado.es_superadmin()) with check (privado.es_superadmin());
-- Cada quien cambia solo su nombre (columna permitida); rol, escuela y activo no.
create policy usuarios_propio on public.usuarios for update to authenticated
  using (id = auth.uid()) with check (id = auth.uid());
revoke update on public.usuarios from authenticated;
grant update (nombre, activo) on public.usuarios to authenticated;
-- 'activo' solo lo cambia quien dirige: un trigger impide que alguien se reactive a sí mismo.
create or replace function privado.proteger_usuario() returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if new.activo is distinct from old.activo and not privado.dirige(old.escuela_id) then
    raise exception 'Solo el director puede activar o desactivar cuentas';
  end if;
  if new.activo = false and old.id = auth.uid() then
    raise exception 'No puedes desactivar tu propia cuenta';
  end if;
  return new;
end $$;
create trigger proteger_usuario before update on public.usuarios
  for each row execute function privado.proteger_usuario();

-- ---------- Invitaciones: el director invita maestros de su escuela; el superadmin, directores ----------
create policy invitaciones_ver on public.invitaciones for select to authenticated
  using (privado.dirige(escuela_id));
create policy invitaciones_crear on public.invitaciones for insert to authenticated
  with check (privado.dirige(escuela_id) and (rol = 'maestro' or privado.es_superadmin()) and usada_en is null);
create policy invitaciones_borrar on public.invitaciones for delete to authenticated
  using (privado.dirige(escuela_id) and usada_en is null);

-- ---------- Pagos: el superadmin los registra; el director ve los de su escuela ----------
create policy pagos_ver on public.pagos for select to authenticated using (privado.dirige(escuela_id));
create policy pagos_superadmin on public.pagos for all to authenticated
  using (privado.es_superadmin()) with check (privado.es_superadmin());

-- ---------- Periodos: los ve la escuela; los administra el director (o el maestro individual) ----------
create policy semestres_ver on public.semestres for select to authenticated
  using (escuela_id = privado.mi_escuela() or privado.es_superadmin());
create policy semestres_admin on public.semestres for all to authenticated
  using (privado.administra(escuela_id)) with check (privado.administra(escuela_id));
create policy parciales_ver on public.parciales for select to authenticated
  using (escuela_id = privado.mi_escuela() or privado.es_superadmin());
create policy parciales_admin on public.parciales for all to authenticated
  using (privado.administra(escuela_id)) with check (privado.administra(escuela_id));

-- ---------- Alumnos, grupos e inscripciones ----------
-- Los ve toda la escuela (para reconocer al mismo alumno entre maestros). Los maestros los crean al
-- importar su Excel y corrigen nombres; borrar solo quien dirige.
create policy alumnos_ver on public.alumnos for select to authenticated
  using (escuela_id = privado.mi_escuela() or privado.es_superadmin());
create policy alumnos_crear on public.alumnos for insert to authenticated
  with check (escuela_id = privado.mi_escuela());
create policy alumnos_editar on public.alumnos for update to authenticated
  using (escuela_id = privado.mi_escuela()) with check (escuela_id = privado.mi_escuela());
create policy alumnos_borrar on public.alumnos for delete to authenticated using (privado.dirige(escuela_id));

create policy grupos_ver on public.grupos for select to authenticated
  using (escuela_id = privado.mi_escuela() or privado.es_superadmin());
create policy grupos_crear on public.grupos for insert to authenticated
  with check (escuela_id = privado.mi_escuela());
create policy grupos_editar on public.grupos for update to authenticated
  using (escuela_id = privado.mi_escuela()) with check (escuela_id = privado.mi_escuela());
create policy grupos_borrar on public.grupos for delete to authenticated using (privado.dirige(escuela_id));

create policy inscripciones_ver on public.inscripciones for select to authenticated
  using (escuela_id = privado.mi_escuela() or privado.es_superadmin());
create policy inscripciones_crear on public.inscripciones for insert to authenticated
  with check (escuela_id = privado.mi_escuela());
create policy inscripciones_editar on public.inscripciones for update to authenticated
  using (escuela_id = privado.mi_escuela()) with check (escuela_id = privado.mi_escuela());
create policy inscripciones_borrar on public.inscripciones for delete to authenticated using (privado.dirige(escuela_id));

-- ---------- Clases: el maestro ve y escribe las suyas; el director ve todas las de su escuela ----------
create policy clases_ver on public.clases for select to authenticated
  using (maestro_id = auth.uid() or privado.dirige(escuela_id));
create policy clases_crear on public.clases for insert to authenticated
  with check (maestro_id = auth.uid() and escuela_id = privado.mi_escuela());
create policy clases_editar on public.clases for update to authenticated
  using (privado.es_mi_clase(id)) with check (maestro_id = auth.uid() and escuela_id = privado.mi_escuela());
create policy clases_borrar on public.clases for delete to authenticated using (privado.es_mi_clase(id));

-- ---------- Lo que cuelga de una clase: lo escribe su maestro, lo lee también quien dirige ----------
create policy secciones_ver on public.secciones for select to authenticated
  using (privado.es_mi_clase(clase_id) or privado.dirige(escuela_id));
create policy secciones_escribir on public.secciones for all to authenticated
  using (privado.es_mi_clase(clase_id)) with check (privado.es_mi_clase(clase_id));

create policy trabajos_ver on public.trabajos for select to authenticated
  using (privado.es_mi_clase(privado.clase_de_seccion(seccion_id)) or privado.dirige(escuela_id));
create policy trabajos_escribir on public.trabajos for all to authenticated
  using (privado.es_mi_clase(privado.clase_de_seccion(seccion_id)))
  with check (privado.es_mi_clase(privado.clase_de_seccion(seccion_id)));

create policy capturas_ver on public.capturas for select to authenticated
  using (privado.es_mi_clase(clase_id) or privado.dirige(escuela_id));
create policy capturas_escribir on public.capturas for all to authenticated
  using (privado.es_mi_clase(clase_id)) with check (privado.es_mi_clase(clase_id));

create policy cuadernillos_ver on public.cuadernillos for select to authenticated
  using (privado.es_mi_clase(clase_id) or privado.dirige(escuela_id));
create policy cuadernillos_escribir on public.cuadernillos for all to authenticated
  using (privado.es_mi_clase(clase_id)) with check (privado.es_mi_clase(clase_id));

-- Cierres: el maestro cierra y reabre (update); nadie los borra desde la app.
create policy cierres_ver on public.cierres for select to authenticated
  using (privado.es_mi_clase(clase_id) or privado.dirige(escuela_id));
create policy cierres_crear on public.cierres for insert to authenticated
  with check (privado.es_mi_clase(clase_id) and cerrado_por = auth.uid());
create policy cierres_reabrir on public.cierres for update to authenticated
  using (privado.es_mi_clase(clase_id)) with check (privado.es_mi_clase(clase_id));
revoke delete on public.cierres from authenticated;
revoke update on public.cierres from authenticated;
grant update (reabierto_por, reabierto_en) on public.cierres to authenticated;

-- Historial: lo leen quien dirige y el maestro sus propios cambios; nadie lo escribe ni borra a mano.
create policy historial_ver on public.historial for select to authenticated
  using (privado.dirige(escuela_id) or usuario_id = auth.uid());
revoke insert, update, delete on public.historial from authenticated;

-- ---------- Triggers: quién y cuándo, e historial ----------
create or replace function privado.sello() returns trigger
language plpgsql set search_path = '' as $$
begin
  new.actualizado_en := now();
  new.actualizado_por := auth.uid();
  return new;
end $$;
create trigger sello before insert or update on public.capturas
  for each row execute function privado.sello();

create or replace function privado.registrar() returns trigger
language plpgsql security definer set search_path = '' as $$
declare r jsonb := case when tg_op = 'DELETE' then to_jsonb(old) else to_jsonb(new) end;
begin
  -- Una captura nueva no se registra (sería duplicar la tabla); sí los cambios y borrados.
  if tg_table_name = 'capturas' and tg_op = 'INSERT' then return new; end if;
  insert into public.historial (escuela_id, tabla, registro_id, accion, antes, despues, usuario_id)
  values ((r ->> 'escuela_id')::uuid, tg_table_name, (r ->> 'id')::uuid, lower(tg_op),
          case when tg_op = 'INSERT' then null else to_jsonb(old) end,
          case when tg_op = 'DELETE' then null else to_jsonb(new) end,
          auth.uid());
  return case when tg_op = 'DELETE' then old else new end;
end $$;
create trigger historial after insert or update or delete on public.capturas
  for each row execute function privado.registrar();
create trigger historial after insert or update or delete on public.cierres
  for each row execute function privado.registrar();
create trigger historial after insert or update or delete on public.secciones
  for each row execute function privado.registrar();
create trigger historial after update or delete on public.alumnos
  for each row execute function privado.registrar();
create trigger historial after update or delete on public.inscripciones
  for each row execute function privado.registrar();

-- ---------- Invitaciones y registro ----------
-- Cuentas que ocupa una escuela (activas + invitaciones pendientes) contra las contratadas.
create or replace function privado.cupo_libre(e uuid) returns int
language sql stable security definer set search_path = '' as $$
  select s.cuentas_contratadas
    - (select count(*) from public.usuarios u where u.escuela_id = e and u.activo and u.rol <> 'superadmin')::int
    - (select count(*) from public.invitaciones i where i.escuela_id = e and i.usada_en is null and i.vence_en > now())::int
  from public.escuelas s where s.id = e
$$;

create or replace function privado.revisar_cupo() returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if privado.cupo_libre(new.escuela_id) < 1 then
    raise exception 'La escuela ya usó todas sus cuentas contratadas';
  end if;
  return new;
end $$;
create trigger revisar_cupo before insert on public.invitaciones
  for each row execute function privado.revisar_cupo();

-- El maestro invitado, ya con su sesión iniciada, acepta con el token del enlace. Un solo uso,
-- con caducidad, y solo para el correo al que se envió.
create or replace function public.aceptar_invitacion(p_token uuid, p_nombre text)
returns public.usuarios
language plpgsql security definer set search_path = '' as $$
declare
  inv public.invitaciones;
  correo_sesion text := lower(auth.jwt() ->> 'email');
  nuevo public.usuarios;
begin
  if auth.uid() is null then raise exception 'Inicia sesión primero'; end if;
  if exists (select 1 from public.usuarios where id = auth.uid()) then
    raise exception 'Esta cuenta ya pertenece a una escuela';
  end if;
  select * into inv from public.invitaciones where token = p_token for update;
  if inv.id is null or inv.usada_en is not null or inv.vence_en < now() then
    raise exception 'La invitación no es válida o ya venció';
  end if;
  if inv.correo <> correo_sesion then
    raise exception 'La invitación es para otro correo';
  end if;
  insert into public.usuarios (id, escuela_id, rol, nombre, correo)
  values (auth.uid(), inv.escuela_id, inv.rol, coalesce(nullif(trim(p_nombre), ''), split_part(correo_sesion, '@', 1)), correo_sesion)
  returning * into nuevo;
  update public.invitaciones set usada_en = now(), usada_por = auth.uid() where id = inv.id;
  return nuevo;
end $$;

-- Paquete individual: el maestro se registra solo y queda como su propia escuela (sin pagar aún:
-- vence en 30 días de prueba; el pago lo marca el superadmin o el cobro automático después).
create or replace function public.registrar_individual(p_nombre text)
returns public.usuarios
language plpgsql security definer set search_path = '' as $$
declare
  correo_sesion text := lower(auth.jwt() ->> 'email');
  esc uuid;
  nuevo public.usuarios;
begin
  if auth.uid() is null then raise exception 'Inicia sesión primero'; end if;
  if exists (select 1 from public.usuarios where id = auth.uid()) then
    raise exception 'Esta cuenta ya pertenece a una escuela';
  end if;
  insert into public.escuelas (nombre, clave, paquete, cuentas_contratadas, vence_el)
  values (coalesce(nullif(trim(p_nombre), ''), correo_sesion),
          'I' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 7)), 'individual', 1, current_date + 30)
  returning id into esc;
  insert into public.usuarios (id, escuela_id, rol, nombre, correo)
  values (auth.uid(), esc, 'maestro', coalesce(nullif(trim(p_nombre), ''), split_part(correo_sesion, '@', 1)), correo_sesion)
  returning * into nuevo;
  return nuevo;
end $$;

revoke execute on function public.aceptar_invitacion(uuid, text) from public, anon;
revoke execute on function public.registrar_individual(text) from public, anon;
grant execute on function public.aceptar_invitacion(uuid, text) to authenticated;
grant execute on function public.registrar_individual(text) to authenticated;
