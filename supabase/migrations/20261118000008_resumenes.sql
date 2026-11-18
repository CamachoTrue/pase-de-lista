-- Resumen de cada clase para el panel del director: lo calcula el celular del maestro (con la misma
-- lógica de la app) y lo sube al sincronizar. El director lo lee sin descargar todas las capturas.
-- datos = { escala, alumnos: { alumno_id: { f, a, F, e, r, p: [..], s } }, secciones: [...] }
create table public.resumenes (
  clase_id uuid primary key,
  escuela_id uuid not null,
  datos jsonb not null,
  ultima_captura date,
  actualizado_en timestamptz not null default now(),
  foreign key (clase_id, escuela_id) references public.clases (id, escuela_id) on delete cascade
);
create index on public.resumenes (escuela_id);
alter table public.resumenes enable row level security;
revoke all on public.resumenes from anon;
grant select, insert, update, delete on public.resumenes to authenticated;
create policy resumenes_ver on public.resumenes for select to authenticated
  using (privado.es_mi_clase(clase_id) or privado.dirige(escuela_id));
create policy resumenes_escribir on public.resumenes for all to authenticated
  using (privado.es_mi_clase(clase_id)) with check (privado.es_mi_clase(clase_id));
