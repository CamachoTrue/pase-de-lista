-- Eliminar mi cuenta (lo exigen Apple y Google a toda app con cuentas).
-- * Maestro independiente (paquete individual): se borra todo lo suyo — grupos, alumnos, capturas,
--   historial y su "escuela" — y su usuario.
-- * Maestro o director de una escuela: se borra su usuario (nombre, correo, contraseña). Las clases y
--   calificaciones son registros de la escuela: se quedan, sin maestro asignado.
-- * El administrador general no se puede borrar desde la app.

-- Las clases ya no exigen maestro: si su cuenta se elimina, la clase queda "sin maestro".
alter table public.clases alter column maestro_id drop not null;
alter table public.clases drop constraint clases_maestro_id_fkey;
alter table public.clases add constraint clases_maestro_id_fkey
  foreign key (maestro_id) references public.usuarios (id) on delete set null;

create or replace function public.eliminar_mi_cuenta()
returns text
language plpgsql security definer set search_path = '' as $$
declare
  yo public.usuarios;
  esc public.escuelas;
begin
  if auth.uid() is null then raise exception 'Inicia sesión primero'; end if;
  select * into yo from public.usuarios where id = auth.uid();
  if yo.rol = 'superadmin' then
    raise exception 'La cuenta de administrador general no se puede eliminar desde la app';
  end if;
  if yo.escuela_id is not null then
    select * into esc from public.escuelas where id = yo.escuela_id;
    if esc.paquete = 'individual'
       and not exists (select 1 from public.usuarios where escuela_id = esc.id and id <> yo.id) then
      delete from public.grupos where escuela_id = esc.id;      -- y en cascada clases, capturas, cierres…
      delete from public.alumnos where escuela_id = esc.id;
      delete from public.historial where escuela_id = esc.id;
      delete from auth.users where id = yo.id;                  -- y en cascada su fila en usuarios
      delete from public.escuelas where id = esc.id;            -- semestres, invitaciones, pagos
      return 'todo';
    end if;
  end if;
  delete from auth.users where id = auth.uid();
  return 'cuenta';
end $$;

revoke all on function public.eliminar_mi_cuenta() from public, anon;
grant execute on function public.eliminar_mi_cuenta() to authenticated;
