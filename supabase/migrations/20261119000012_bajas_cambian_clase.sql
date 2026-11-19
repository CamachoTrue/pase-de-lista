-- Dar de baja (o reactivar) a un alumno cambia su inscripción: las clases de ese grupo se marcan como
-- cambiadas por este dispositivo, para que los demás dispositivos del maestro lo bajen.
create or replace function privado.cambio_en_inscripcion() returns trigger
language plpgsql security definer set search_path = '' as $$
declare
  c uuid;
begin
  for c in select id from public.clases where grupo_id = coalesce(new.grupo_id, old.grupo_id) loop
    perform privado.marcar_clase(c);
  end loop;
  return null;
end $$;

create trigger cambio_en_inscripcion after update on public.inscripciones
  for each row execute function privado.cambio_en_inscripcion();
