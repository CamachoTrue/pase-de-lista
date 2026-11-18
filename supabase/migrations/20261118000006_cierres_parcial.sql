-- Cierres desde el celular: el parcial se guarda por número (1–6) y el nombre del semestre, porque el
-- maestro puede tener sus propias fechas (paquete individual o escuela que aún no las define).
alter table public.cierres add column parcial int check (parcial between 1 and 6);
alter table public.cierres add column semestre text not null default '';
alter table public.cierres drop constraint if exists cierres_check;
alter table public.cierres add constraint cierres_tipo_parcial
  check ((tipo = 'parcial') = (parcial is not null or parcial_id is not null));
