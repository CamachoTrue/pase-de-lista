-- Ver los cambios de otro dispositivo: cada clase guarda cuándo la cambió por última vez cada
-- dispositivo ({ "<dispositivo>": "<fecha y hora>" }). Un celular revisa esa lista cada 30 s y, si otro
-- dispositivo cambió la clase después de lo que ya tiene, baja la clase. Sus propios cambios no cuentan.
-- El dispositivo llega en el encabezado x-dispositivo de cada petición (lo pone la app).
alter table public.clases add column cambios jsonb not null default '{}'::jsonb;

create or replace function privado.dispositivo() returns text
language sql stable set search_path = '' as $$
  select coalesce(nullif(left(nullif(current_setting('request.headers', true), '')::json ->> 'x-dispositivo', 64), ''), 'otro')
$$;

-- Marca la clase como cambiada por este dispositivo (una sola escritura por transacción: now() no cambia).
create or replace function privado.marcar_clase(c uuid) returns void
language sql security definer set search_path = '' as $$
  update public.clases
     set cambios = cambios || jsonb_build_object(privado.dispositivo(), now())
   where id = c and (cambios ->> privado.dispositivo()) is distinct from to_jsonb(now()) #>> '{}'
$$;

create or replace function privado.cambio_en_clase() returns trigger
language plpgsql security definer set search_path = '' as $$
declare
  fila record := coalesce(new, old);
begin
  if tg_table_name = 'trabajos' then
    perform privado.marcar_clase((select clase_id from public.secciones where id = fila.seccion_id));
  else
    perform privado.marcar_clase(fila.clase_id);
  end if;
  return null;
end $$;

create trigger cambio_en_clase after insert or update or delete on public.capturas
  for each row execute function privado.cambio_en_clase();
create trigger cambio_en_clase after insert or update or delete on public.secciones
  for each row execute function privado.cambio_en_clase();
create trigger cambio_en_clase after insert or update or delete on public.trabajos
  for each row execute function privado.cambio_en_clase();
create trigger cambio_en_clase after insert or update or delete on public.cierres
  for each row execute function privado.cambio_en_clase();
create trigger cambio_en_clase after insert or update or delete on public.cuadernillos
  for each row execute function privado.cambio_en_clase();

-- Cambios a la clase misma (materia, escala, retardo, calificación semestral).
create or replace function privado.cambio_de_clase() returns trigger
language plpgsql set search_path = '' as $$
begin
  if (new.materia, new.retardo, new.escala, new.config_semestral)
     is distinct from (old.materia, old.retardo, old.escala, old.config_semestral) then
    new.cambios := new.cambios || jsonb_build_object(privado.dispositivo(), now());
  end if;
  return new;
end $$;
create trigger cambio_de_clase before update on public.clases
  for each row execute function privado.cambio_de_clase();

revoke all on function privado.marcar_clase(uuid) from public;
