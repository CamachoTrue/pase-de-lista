-- Al cerrar un parcial se guardan también los criterios con que se cerró (porcentajes de cada sección,
-- retardo y escala). Así ese parcial se exporta igual aunque el siguiente cambie de secciones o porcentajes.
alter table public.cierres add column criterios jsonb;
