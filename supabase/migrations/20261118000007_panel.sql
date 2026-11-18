-- Panel del superadmin: precios de lista, escuelas con su cobro y pagos. Y el estado de la suscripción
-- (activa / vence) en el perfil, para que la app avise y deje de subir cuando venció.

-- Precios de lista (MXN). Una escuela con precio_mensual propio usa ese (trato especial).
create table public.precios_lista (
  paquete public.paquete primary key,
  por_mes numeric(10, 2) not null check (por_mes >= 0),     -- por maestro
  por_anio numeric(10, 2) not null check (por_anio >= 0),   -- por maestro
  dias_prueba int not null default 30 check (dias_prueba >= 0),
  actualizado_en timestamptz not null default now()
);
insert into public.precios_lista (paquete, por_mes, por_anio, dias_prueba) values
  ('individual', 99, 899, 30),
  ('empresarial', 69, 690, 0);
alter table public.precios_lista enable row level security;
revoke all on public.precios_lista from anon;
grant select, insert, update, delete on public.precios_lista to authenticated;
create policy precios_ver on public.precios_lista for select to authenticated using (true);
create policy precios_superadmin on public.precios_lista for all to authenticated
  using (privado.es_superadmin()) with check (privado.es_superadmin());

-- Todas las escuelas con lo que se les cobra (solo superadmin).
create or replace function public.panel_escuelas()
returns table (id uuid, nombre text, clave text, paquete public.paquete, activa boolean, precio_mensual numeric,
               precio_lista numeric, cuentas_contratadas int, cuentas_usadas int, invitaciones int, vence_el date,
               ultimo_pago date, creado_en timestamptz)
language plpgsql stable security definer set search_path = '' as $$
begin
  if not privado.es_superadmin() then raise exception 'Solo el administrador general'; end if;
  return query
  select s.id, s.nombre, s.clave, s.paquete, s.activa, s.precio_mensual, pl.por_mes, s.cuentas_contratadas,
    (select count(*) from public.usuarios u where u.escuela_id = s.id and u.activo and u.rol <> 'superadmin')::int,
    (select count(*) from public.invitaciones i where i.escuela_id = s.id and i.usada_en is null and i.vence_en > now())::int,
    s.vence_el,
    (select max(p.periodo_fin) from public.pagos p where p.escuela_id = s.id),
    s.creado_en
  from public.escuelas s left join public.precios_lista pl on pl.paquete = s.paquete
  order by s.creado_en desc;
end $$;
revoke execute on function public.panel_escuelas() from public, anon;
grant execute on function public.panel_escuelas() to authenticated;

-- Registrar un pago (transferencia, efectivo…): queda el pago y la escuela se extiende hasta el fin del periodo.
create or replace function public.registrar_pago(p_escuela uuid, p_monto numeric, p_inicio date, p_fin date, p_metodo text, p_factura text)
returns void
language plpgsql security definer set search_path = '' as $$
begin
  if not privado.es_superadmin() then raise exception 'Solo el administrador general'; end if;
  insert into public.pagos (escuela_id, periodo_inicio, periodo_fin, monto, metodo, factura, registrado_por)
  values (p_escuela, p_inicio, p_fin, p_monto, coalesce(nullif(trim(p_metodo), ''), 'transferencia'), nullif(trim(p_factura), ''), auth.uid());
  update public.escuelas set vence_el = greatest(coalesce(vence_el, p_fin), p_fin), activa = true where id = p_escuela;
end $$;
revoke execute on function public.registrar_pago(uuid, numeric, date, date, text, text) from public, anon;
grant execute on function public.registrar_pago(uuid, numeric, date, date, text, text) to authenticated;

-- El perfil ahora dice si la escuela está activa y cuándo vence.
drop function public.mi_perfil();
create function public.mi_perfil()
returns table (nombre text, correo text, rol public.rol, escuela_id uuid, escuela text, clave text, paquete public.paquete,
               activa boolean, vence_el date)
language sql stable security definer set search_path = '' as $$
  select u.nombre, u.correo, u.rol, u.escuela_id, s.nombre, s.clave, s.paquete, s.activa, s.vence_el
  from public.usuarios u left join public.escuelas s on s.id = u.escuela_id
  where u.id = auth.uid() and u.activo
$$;
revoke execute on function public.mi_perfil() from public, anon;
grant execute on function public.mi_perfil() to authenticated;

-- El paquete individual usa los días de prueba de la lista de precios.
create or replace function public.registrar_individual(p_nombre text)
returns public.usuarios
language plpgsql security definer set search_path = '' as $$
declare
  correo_sesion text := lower(auth.jwt() ->> 'email');
  esc uuid;
  dias int := coalesce((select dias_prueba from public.precios_lista where paquete = 'individual'), 30);
  nuevo public.usuarios;
begin
  if auth.uid() is null then raise exception 'Inicia sesión primero'; end if;
  if exists (select 1 from public.usuarios where id = auth.uid()) then
    raise exception 'Esta cuenta ya pertenece a una escuela';
  end if;
  insert into public.escuelas (nombre, clave, paquete, cuentas_contratadas, vence_el)
  values (coalesce(nullif(trim(p_nombre), ''), correo_sesion),
          'I' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 7)), 'individual', 1, current_date + dias)
  returning id into esc;
  insert into public.usuarios (id, escuela_id, rol, nombre, correo)
  values (auth.uid(), esc, 'maestro', coalesce(nullif(trim(p_nombre), ''), split_part(correo_sesion, '@', 1)), correo_sesion)
  returning * into nuevo;
  return nuevo;
end $$;
