-- Al abrir el enlace de invitación (todavía sin cuenta) la app muestra a qué escuela y con qué correo
-- invitaron a la persona. Solo con el token exacto (imposible de adivinar); no revela nada más.
create or replace function public.ver_invitacion(p_token uuid)
returns table (escuela text, correo text, rol public.rol, valida boolean)
language sql stable security definer set search_path = '' as $$
  select s.nombre, i.correo, i.rol, (i.usada_en is null and i.vence_en > now() and s.activa)
  from public.invitaciones i join public.escuelas s on s.id = i.escuela_id
  where i.token = p_token
$$;
revoke execute on function public.ver_invitacion(uuid) from public;
grant execute on function public.ver_invitacion(uuid) to anon, authenticated;

-- Perfil de quien inició sesión, con el nombre de su escuela (para la pantalla Mi cuenta).
create or replace function public.mi_perfil()
returns table (nombre text, correo text, rol public.rol, escuela_id uuid, escuela text, clave text, paquete public.paquete)
language sql stable security definer set search_path = '' as $$
  select u.nombre, u.correo, u.rol, u.escuela_id, s.nombre, s.clave, s.paquete
  from public.usuarios u left join public.escuelas s on s.id = u.escuela_id
  where u.id = auth.uid() and u.activo
$$;
revoke execute on function public.mi_perfil() from public, anon;
grant execute on function public.mi_perfil() to authenticated;
