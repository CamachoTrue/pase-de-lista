# Base de datos (Supabase)

Proyecto: `nwgswikzkuuhhisdqbjq` (región West US). Las tablas y permisos viven en `migrations/` y se aplican
en orden; nunca se cambian a mano en el panel.

| Migración | Qué hace |
|---|---|
| `…0001_esquema.sql` | 16 tablas: escuelas, cuentas, periodos, alumnos, grupos, clases, secciones, capturas, cierres, historial… |
| `…0002_permisos.sql` | Row Level Security por escuela y rol, historial automático, invitaciones y registro individual |
| `…0003_invitacion_publica.sql` | `ver_invitacion(token)` para el enlace (sin sesión) y `mi_perfil()` |

## Probar en local (sin tocar el proyecto real)
`npm test` corre las migraciones en un Postgres local (PGlite) y prueba los permisos con dos escuelas
ficticias (`tests/rls.test.js`).

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
