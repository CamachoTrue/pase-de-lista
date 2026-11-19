-- Un grupo con dos clases del mismo maestro (una sin materia y otra con materia) partía lo capturado en
-- dos. Ahora importar_grupo reutiliza la clase que el maestro ya tiene en el grupo, una materia vacía no
-- borra la de la clase, y guardar_secciones reconoce por nombre las secciones de un dispositivo que se
-- volvió a ligar a otra clase.
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
  mat text;
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
  -- La clase del maestro en este grupo: la de la misma materia. Si no la hay, se reutiliza la que ya
  -- tiene (en vez de crear otra y partir lo capturado en dos): la que no tenía materia adopta la nueva,
  -- y un dispositivo sin materia se liga a la única clase que el maestro tiene en el grupo.
  mat := coalesce(trim(p_materia), '');
  select id into c from public.clases
    where grupo_id = g and maestro_id = auth.uid() and materia = mat;
  if c is null and mat <> '' then
    select id into c from public.clases
      where grupo_id = g and maestro_id = auth.uid() and materia = '' order by creado_en limit 1;
    if c is not null then update public.clases set materia = mat where id = c; end if;
  end if;
  if c is null and mat = '' and (select count(*) from public.clases where grupo_id = g and maestro_id = auth.uid()) = 1 then
    select id into c from public.clases where grupo_id = g and maestro_id = auth.uid();
  end if;
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
  -- Una materia vacía no borra la que ya tiene la clase (otro dispositivo pudo no tenerla puesta).
  update public.clases set materia = coalesce(nullif(trim(p_materia), ''), materia), retardo = p_retardo, escala = p_escala
    where id = p_clase returning escuela_id into esc;
  if esc is null then raise exception 'Clase no encontrada'; end if;
  -- Primero se apartan los nombres (para permitir intercambios de nombre entre secciones).
  update public.secciones set nombre = nombre || ' ~' || left(id::text, 8)
    where clase_id = p_clase and id in (select (x ->> 'id')::uuid from jsonb_array_elements(p_secciones) x where x ->> 'id' is not null);
  for s in select * from jsonb_array_elements(p_secciones) loop
    sid := nullif(s ->> 'id', '')::uuid;
    -- Sección que viene de otra clase (dispositivo que se volvió a ligar): se reconoce por su nombre.
    -- (Mismo nombre y plantilla; nunca una que venga en la lista con su propio id.)
    if sid is null or not exists (select 1 from public.secciones where id = sid and clase_id = p_clase) then
      select x.id into sid from public.secciones x
        where x.clase_id = p_clase and x.nombre = s ->> 'nombre' and x.plantilla = (s ->> 'plantilla')::public.plantilla
          and not (x.id = any (ids)) limit 1;
    end if;
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
