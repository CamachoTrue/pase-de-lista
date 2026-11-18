-- Sincronización desde el celular. Todas son "security invoker": corren con los permisos de quien
-- llama, así que las reglas de RLS siguen aplicando (un maestro solo toca sus clases).

-- Nombre comparable: sin acentos, mayúsculas y espacios sencillos.
create or replace function privado.norm(t text) returns text
language sql immutable set search_path = '' as $$
  select upper(translate(regexp_replace(trim(coalesce(t, '')), '\s+', ' ', 'g'),
    'áéíóúüñÁÉÍÓÚÜÑàèìòùÀÈÌÒÙ', 'aeiouunAEIOUUNaeiouAEIOU'))
$$;
grant execute on function privado.norm(text) to authenticated;

-- Importar un grupo del Excel del maestro: crea el grupo si no existe, reconoce a los alumnos que
-- ya están inscritos (mismo nombre) y crea a los nuevos con su clave ESC01-V205-001; luego crea
-- (o reutiliza) la clase del maestro. Devuelve los ids en el mismo orden en que llegaron los alumnos.
create or replace function public.importar_grupo(p_grupo text, p_materia text, p_alumnos jsonb)
returns jsonb
language plpgsql security invoker set search_path = '' as $$
declare
  esc uuid := privado.mi_escuela();
  clave_esc text;
  grupo_clave text := regexp_replace(upper(coalesce(p_grupo, '')), '[^A-Z0-9]', '', 'g');
  g uuid;
  c uuid;
  a jsonb;
  aid uuid;
  num text;
  nueva text;
  k int;
  res jsonb := '[]'::jsonb;
begin
  if esc is null then raise exception 'Tu cuenta no pertenece a una escuela'; end if;
  if length(trim(coalesce(p_grupo, ''))) = 0 then raise exception 'Falta el nombre del grupo'; end if;
  select s.clave into clave_esc from public.escuelas s where s.id = esc;
  select id into g from public.grupos where escuela_id = esc and semestre_id is null and nombre = trim(p_grupo);
  if g is null then
    insert into public.grupos (escuela_id, nombre) values (esc, trim(p_grupo)) returning id into g;
  end if;
  for a in select * from jsonb_array_elements(p_alumnos) loop
    num := coalesce(nullif(trim(a ->> 'numero'), ''), '0');
    select al.id into aid
      from public.inscripciones i join public.alumnos al on al.id = i.alumno_id
      where i.grupo_id = g and privado.norm(al.nombre) = privado.norm(a ->> 'nombre')
      limit 1;
    if aid is null then
      nueva := clave_esc || '-' || grupo_clave || '-' || lpad(regexp_replace(num, '[^0-9]', '', 'g'), 3, '0');
      k := 1;
      while exists (select 1 from public.alumnos al where al.escuela_id = esc and al.clave = nueva || case when k > 1 then '-' || k else '' end) loop
        k := k + 1;
      end loop;
      if k > 1 then nueva := nueva || '-' || k; end if;
      insert into public.alumnos (escuela_id, clave, nombre) values (esc, nueva, trim(a ->> 'nombre')) returning id into aid;
      insert into public.inscripciones (escuela_id, grupo_id, alumno_id, numero_lista) values (esc, g, aid, num);
    end if;
    res := res || jsonb_build_array(aid);
  end loop;
  select id into c from public.clases
    where grupo_id = g and maestro_id = auth.uid() and materia = coalesce(trim(p_materia), '');
  if c is null then
    insert into public.clases (escuela_id, grupo_id, maestro_id, materia)
      values (esc, g, auth.uid(), coalesce(trim(p_materia), '')) returning id into c;
  end if;
  return jsonb_build_object('grupo_id', g, 'clase_id', c, 'alumnos', res);
end $$;

-- Secciones y calificación de una clase tal como están en el celular. Las que ya no vienen se borran
-- (con sus capturas). Devuelve { id_local: id_nube }.
create or replace function public.guardar_secciones(p_clase uuid, p_materia text, p_retardo int, p_escala int, p_secciones jsonb)
returns jsonb
language plpgsql security invoker set search_path = '' as $$
declare
  esc uuid;
  s jsonb;
  sid uuid;
  ids uuid[] := '{}';
  res jsonb := '{}'::jsonb;
begin
  update public.clases set materia = coalesce(trim(p_materia), ''), retardo = p_retardo, escala = p_escala
    where id = p_clase returning escuela_id into esc;
  if esc is null then raise exception 'Clase no encontrada'; end if;
  -- Primero se apartan los nombres (para permitir intercambios de nombre entre secciones).
  update public.secciones set nombre = nombre || ' ~' || left(id::text, 8)
    where clase_id = p_clase and id in (select (x ->> 'id')::uuid from jsonb_array_elements(p_secciones) x where x ->> 'id' is not null);
  for s in select * from jsonb_array_elements(p_secciones) loop
    sid := nullif(s ->> 'id', '')::uuid;
    if sid is not null and exists (select 1 from public.secciones where id = sid and clase_id = p_clase) then
      update public.secciones set nombre = s ->> 'nombre', porcentaje = coalesce((s ->> 'porcentaje')::numeric, 0),
        minimo = (s ->> 'minimo')::int, orden = (s ->> 'orden')::int
        where id = sid and (nombre, porcentaje, minimo, orden) is distinct from
          (s ->> 'nombre', coalesce((s ->> 'porcentaje')::numeric, 0), (s ->> 'minimo')::int, (s ->> 'orden')::int);
    else
      insert into public.secciones (escuela_id, clase_id, nombre, plantilla, porcentaje, minimo, orden)
        values (esc, p_clase, s ->> 'nombre', (s ->> 'plantilla')::public.plantilla, coalesce((s ->> 'porcentaje')::numeric, 0),
                (s ->> 'minimo')::int, (s ->> 'orden')::int)
        returning id into sid;
    end if;
    ids := ids || sid;
    res := res || jsonb_build_object(s ->> 'local', sid);
  end loop;
  delete from public.secciones where clase_id = p_clase and not (id = any (ids));
  return res;
end $$;

-- Trabajos (plantilla 0/50/100) de varios días: deja en cada (sección, fecha) exactamente los trabajos
-- que vienen, en orden. Devuelve [{ seccion, fecha, ids: [...] }].
create or replace function public.guardar_trabajos(p_dias jsonb)
returns jsonb
language plpgsql security invoker set search_path = '' as $$
declare
  d jsonb;
  t jsonb;
  tid uuid;
  ids uuid[];
  esc uuid;
  res jsonb := '[]'::jsonb;
  i int;
begin
  for d in select * from jsonb_array_elements(p_dias) loop
    select escuela_id into esc from public.secciones where id = (d ->> 'seccion')::uuid;
    ids := '{}';
    i := 0;
    for t in select * from jsonb_array_elements(d -> 'trabajos') loop
      tid := nullif(t ->> 'id', '')::uuid;
      if tid is not null and exists (select 1 from public.trabajos where id = tid and seccion_id = (d ->> 'seccion')::uuid) then
        update public.trabajos set titulo = coalesce(t ->> 'titulo', ''), orden = i
          where id = tid and (titulo, orden) is distinct from (coalesce(t ->> 'titulo', ''), i);
      else
        insert into public.trabajos (escuela_id, seccion_id, fecha, titulo, orden)
          values (esc, (d ->> 'seccion')::uuid, (d ->> 'fecha')::date, coalesce(t ->> 'titulo', ''), i) returning id into tid;
      end if;
      ids := ids || tid;
      i := i + 1;
    end loop;
    delete from public.trabajos
      where seccion_id = (d ->> 'seccion')::uuid and fecha = (d ->> 'fecha')::date and not (id = any (ids));
    res := res || jsonb_build_object('seccion', d ->> 'seccion', 'fecha', d ->> 'fecha', 'ids', to_jsonb(ids));
  end loop;
  return res;
end $$;

-- Capturas de varios días: [{ seccion, fecha, trabajo?, valores: { alumno_id: valor } }]. Deja el día
-- exactamente como viene: inserta lo nuevo, cambia solo lo que cambió (así el historial registra
-- cambios reales) y borra lo que ya no está.
create or replace function public.guardar_dias(p_dias jsonb)
returns int
language plpgsql security invoker set search_path = '' as $$
declare
  d jsonb;
  n int := 0;
begin
  for d in select * from jsonb_array_elements(p_dias) loop
    insert into public.capturas (escuela_id, clase_id, seccion_id, alumno_id, fecha, trabajo_id, valor)
      select s.escuela_id, s.clase_id, s.id, kv.key::uuid, (d ->> 'fecha')::date, nullif(d ->> 'trabajo', '')::uuid, kv.value
      from public.secciones s, jsonb_each_text(coalesce(d -> 'valores', '{}'::jsonb)) kv
      where s.id = (d ->> 'seccion')::uuid
    on conflict (seccion_id, alumno_id, fecha, trabajo_id)
      do update set valor = excluded.valor where public.capturas.valor is distinct from excluded.valor;
    delete from public.capturas c
      where c.seccion_id = (d ->> 'seccion')::uuid and c.fecha = (d ->> 'fecha')::date
        and c.trabajo_id is not distinct from nullif(d ->> 'trabajo', '')::uuid
        and not (coalesce(d -> 'valores', '{}'::jsonb) ? c.alumno_id::text);
    n := n + 1;
  end loop;
  return n;
end $$;

revoke execute on function public.importar_grupo(text, text, jsonb) from public, anon;
revoke execute on function public.guardar_secciones(uuid, text, int, int, jsonb) from public, anon;
revoke execute on function public.guardar_trabajos(jsonb) from public, anon;
revoke execute on function public.guardar_dias(jsonb) from public, anon;
grant execute on function public.importar_grupo(text, text, jsonb) to authenticated;
grant execute on function public.guardar_secciones(uuid, text, int, int, jsonb) to authenticated;
grant execute on function public.guardar_trabajos(jsonb) to authenticated;
grant execute on function public.guardar_dias(jsonb) to authenticated;
