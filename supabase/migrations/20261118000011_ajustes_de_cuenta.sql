-- Ajustes de la cuenta que antes vivían solo en el celular: "Valores de cada estado" (A, F, R, J, Sí, No)
-- y los parciales que el maestro define él mismo (cuando su escuela no los pone). Así salen iguales en
-- todos sus dispositivos. ajustes_en la pone la base de datos: es la hora del último cambio.
alter table public.usuarios
  add column ajustes jsonb not null default '{}'::jsonb,
  add column ajustes_en timestamptz;

grant update (ajustes) on public.usuarios to authenticated;

create or replace function privado.sello_ajustes() returns trigger
language plpgsql security definer set search_path = '' as $$
begin
  if new.ajustes is distinct from old.ajustes then
    -- Solo cada quien cambia sus propios ajustes (el director no cambia los de sus maestros).
    if old.id is distinct from auth.uid() then
      raise exception 'Solo puedes cambiar tus propios ajustes';
    end if;
    new.ajustes_en := clock_timestamp();
  else
    new.ajustes_en := old.ajustes_en;
  end if;
  return new;
end $$;
create trigger sello_ajustes before update on public.usuarios
  for each row execute function privado.sello_ajustes();
