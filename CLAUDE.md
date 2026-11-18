# Pase de Lista — reglas del proyecto

App web instalable (PWA) para que maestros de preparatoria pasen lista y califiquen desde el celular.
Publicada en https://pase-de-lista-omega.vercel.app (Vercel publica solo cada push a `main`).
Plan del paso 1 (Supabase y login): ver el documento "Pase de Lista — Plan paso 1" del dueño.

## Quién la usa
Maestros en el aula, con el celular en una mano, a veces sin internet, con prisa. Todo se diseña para
pantalla de 375 px, toques grandes, textos en español claro y sin tecnicismos.

## Comandos
- `npm run dev` — app en http://localhost:5173
- `npm test` — pruebas de `src/lib/excel-core.js` (lista ficticia)
- `npm run check` — pruebas + build + presupuesto de peso. **Debe pasar antes de cualquier commit.**

## Estructura
- `src/main.js` arranque y migración de datos guardados · `src/app/*` un módulo por función
- `src/lib/excel-core.js` lógica pura (Excel, secciones, calificaciones): sin DOM, con pruebas
- `public/` manifest, `sw.js` (el build le escribe archivos y versión), íconos, librerías en `vendor/`

## Reglas que no se rompen
1. **Nunca publicar sin permiso.** Push a `main` = publicar al maestro. Trabajar en una rama, probar,
   y preguntar al dueño antes de mezclar a `main`. Nada de `--force`.
2. **Los datos del maestro son sagrados.** Están en su celular (localStorage `paseDeLista.v1` +
   IndexedDB para PDFs). Cualquier cambio de formato lleva migración en `src/main.js` que conserve
   todo, y una prueba. Nunca borrar ni reescribir datos guardados sin migrarlos.
3. **Pruebas en el navegador sin tocar datos reales:** usar la fecha libre 2027-01-20, confirmar
   antes que esté vacía, y borrar lo de prueba al terminar. Regresar grupo, fecha y pestaña como estaban.
4. **Sin datos reales de alumnos en el repositorio** (nombres, listas, Excel). Las pruebas usan
   listas ficticias. `*.xlsx` está en `.gitignore`.
5. **Rápida y ligera** (presupuesto en `scripts/check-size.mjs`): al abrir, máximo 55 KB comprimidos.
   Librerías pesadas (ExcelJS, PDF.js) se cargan solo al usarse. Nada de frameworks ni dependencias
   nuevas sin justificar el peso y pedir permiso.
6. **Funciona sin internet.** Toda función nueva debe servir offline o decir claramente que necesita
   conexión. Las capturas nunca dependen de la red.
7. **Calificaciones:** la lógica vive solo en `excel-core.js` y toda regla nueva lleva prueba.
   La calificación empieza completa y solo descuenta lo perdido (secciones sin datos no descuentan).
8. **Interfaz en español**, sin anglicismos innecesarios; confirmaciones con la app (`confirmBulk`,
   `confirmChange`), nunca `confirm()`/`alert()` del navegador (se bloquean en el celular).
9. **Seguridad (Supabase):** permisos con Row Level Security en la base de datos; la clave
   `service_role` y contraseñas jamás en el código, el repo ni el chat.

## Equipo de agentes (`.claude/agents/`)
| Agente | Cuándo |
|---|---|
| `frontend` | Pantallas, módulos de `src/app`, estilos |
| `backend` | Supabase: tablas, migraciones SQL, RLS, login, sincronización |
| `disenador` | Revisar o proponer diseño, textos y flujo antes de construir |
| `qa` | Después de cada cambio: pruebas, build, peso, revisión en navegador a 375 px |
| `rendimiento` | Cuando algo nuevo pesa o se siente lento; antes de agregar librerías |
| `seguridad` | Cambios de login, permisos, datos personales o secretos |

Flujo normal de una función: `disenador` (si cambia la interfaz) → `frontend`/`backend` → `qa`
(+ `rendimiento` y `seguridad` cuando aplique) → el dueño aprueba → se mezcla a `main`.
