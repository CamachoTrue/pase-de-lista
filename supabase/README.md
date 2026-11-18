# Base de datos (Supabase)

Proyecto: `nwgswikzkuuhhisdqbjq` (región West US). Las tablas y permisos viven en `migrations/` y se aplican
en orden; nunca se cambian a mano en el panel.

| Migración | Qué hace |
|---|---|
| `…0001_esquema.sql` | 16 tablas: escuelas, cuentas, periodos, alumnos, grupos, clases, secciones, capturas, cierres, historial… |
| `…0002_permisos.sql` | Row Level Security por escuela y rol, historial automático, invitaciones y registro individual |
| `…0003_invitacion_publica.sql` | `ver_invitacion(token)` para el enlace (sin sesión) y `mi_perfil()` |
| `…0005_archivos.sql` | Bucket privado `archivos`: `<escuela>/<usuario>/excel/…` y `/cuadernillos/<clase>.pdf`, cada quien su carpeta |
| `…0006_cierres_parcial.sql` | Cierres con número de parcial y nombre de semestre (fechas propias del maestro) |
| `…0004_sincronizar.sql` | `importar_grupo`, `guardar_secciones`, `guardar_trabajos`, `guardar_dias` (security invoker: aplica RLS) |

## Probar en local (sin tocar el proyecto real)
`npm test` corre las migraciones en un Postgres local (PGlite) y prueba los permisos con dos escuelas
ficticias (`tests/rls.test.js`).

## Sincronización (app → nube)
`src/app/sync.js` sube solo lo que cambió (`src/lib/sync-plan.js`, con pruebas): por cada (sección, fecha) guarda una
huella de lo subido; si cambia, manda el día completo y `guardar_dias` lo deja igual (solo cambia lo distinto, así el
historial no se llena de repetidos). Corre unos segundos después de cada cambio, al volver el internet y al abrir.
También sube el Excel original y los cuadernillos (Storage). En un dispositivo nuevo, Mi cuenta ofrece
"Bajar a este celular" (`src/app/bajar.js` + `armarGrupo` en `sync-plan.js`): baja Excel, clases, secciones, alumnos,
capturas y cuadernillos, y deja todo marcado como ya subido.

> `supabase db push/query` desde la consola interna de Claude se cuelga: correrlos en la terminal.

## Aplicar al proyecto
```bash
npx supabase login                                   # una vez; abre el navegador
npx supabase link --project-ref nwgswikzkuuhhisdqbjq
npx supabase db push                                 # aplica las migraciones nuevas
```
Nunca escribir aquí la contraseña de la base de datos ni la clave `service_role`.

## Login (config.toml → `npx supabase config push`)
- Correo y contraseña (mínimo 8). Sin correo de confirmación: el servicio de correo gratis de Supabase solo
  manda a miembros del equipo, así que la cuenta se crea desde el enlace de invitación que manda el director.
- La sesión se guarda en el celular ("Mantener la sesión iniciada") y se renueva sola.
- Antes de `config push`, revisar `npx supabase config diff`: el archivo debe coincidir con el proyecto salvo
  lo que se quiere cambiar.
- Pendiente: correo propio (SMTP, p. ej. Resend) para "olvidé mi contraseña" y confirmar correos.

## Pendiente para etapas siguientes
- Permisos de Storage para los PDFs de cuadernillos.
- Panel de superadmin: función para leer precios y cuentas de todas las escuelas (hoy `datos_cobro()` da solo la propia).
