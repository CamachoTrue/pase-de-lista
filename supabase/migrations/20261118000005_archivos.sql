-- Archivos de cada maestro en Supabase Storage (privado): su Excel original (la app lo necesita para
-- exportar) y los cuadernillos PDF de sus clases. Ruta: <escuela_id>/<usuario_id>/...
-- Cada quien solo lee y escribe dentro de su propia carpeta.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('archivos', 'archivos', false, 52428800,
        array['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/octet-stream'])
on conflict (id) do nothing;

create policy archivos_propios_ver on storage.objects for select to authenticated
  using (bucket_id = 'archivos'
         and (storage.foldername(name))[1] = privado.mi_escuela()::text
         and (storage.foldername(name))[2] = auth.uid()::text);
create policy archivos_propios_subir on storage.objects for insert to authenticated
  with check (bucket_id = 'archivos'
              and (storage.foldername(name))[1] = privado.mi_escuela()::text
              and (storage.foldername(name))[2] = auth.uid()::text);
create policy archivos_propios_cambiar on storage.objects for update to authenticated
  using (bucket_id = 'archivos'
         and (storage.foldername(name))[1] = privado.mi_escuela()::text
         and (storage.foldername(name))[2] = auth.uid()::text)
  with check (bucket_id = 'archivos'
              and (storage.foldername(name))[1] = privado.mi_escuela()::text
              and (storage.foldername(name))[2] = auth.uid()::text);
create policy archivos_propios_borrar on storage.objects for delete to authenticated
  using (bucket_id = 'archivos'
         and (storage.foldername(name))[1] = privado.mi_escuela()::text
         and (storage.foldername(name))[2] = auth.uid()::text);
