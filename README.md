# Pase de Lista

App web instalable (PWA) para pasar lista desde un Excel. Funciona en computadora y en celular.

**App publicada:** https://pase-de-lista-omega.vercel.app (cada push a `main` se publica solo en Vercel)

## Funciones
- **Subir Excel**: cada hoja es un grupo; la columna **NOMBRE** se detecta sola.
- **Pase de lista uno por uno**: un alumno en pantalla con 4 opciones —
  **A** Asistencia, **F** Falta, **R** Retardo, **J** Justificado. Al elegir, pasa al siguiente.
  Empieza en el primer alumno sin marcar; botones Anterior / Saltar.
- **Pestañas por sección**: Sí/No y 0/50/100 se capturan
  en una lista compacta (unos 13 alumnos por pantalla en celular). Al final, botones para poner Sí o No (Lecturas) o 0 (Trabajos) a los que quedaron sin marcar. Si en Lecturas todos tienen la misma marca, un botón los cambia todos a la otra.
- **Varios trabajos por día**: en Trabajos, flechas para cambiar de trabajo, botón "+ Otro trabajo" y un nombre opcional
  para cada uno. En el Excel cada trabajo es una columna "dd/mm/aa Trabajo: nombre".
- **Disciplina**: todos tienen 10 sin capturar nada; el maestro busca al alumno y toca 0 si se portó mal.
  En el Excel no hay una columna por día: al final van 2 columnas, *Días con 0* y *Fechas*.
- **Examen**: se escribe el porcentaje (0 a 100) de cada alumno; Enter pasa al siguiente.
  En el Excel solo hay columna en las fechas con examen, al final de todo.
- **Materia de cada grupo** (⋮ → Materia de este grupo): para cuando el mismo grupo viene dos veces en el Excel
  (una hoja por materia). El selector muestra p. ej. "V205 · Literatura"; también sale en el resumen por alumno.
- **Secciones y calificaciones** (⋮ → Secciones y calificaciones): cada grupo tiene sus propias secciones. El maestro
  puede **agregar, quitar y renombrar** secciones y elegir la **plantilla** de cada una: Sí/No (con mínimo de "Sí" para el 100 %),
  0/50/100 (varios por día, con título), Disciplina (todos 10, marcar 0) o Calificación de 0 a 100 (como examen).
  Asistencia (A/F/R/J) siempre está. Por defecto: Asistencia, Lecturas, Trabajos, Disciplina y Examen.
  Cada sección es una pestaña y lleva su porcentaje; la calificación se calcula sola (las secciones sin datos aún no cuentan).
  En el Excel van al final, por sección y total (sobre 10 o 100). Quitar una sección con datos pide confirmación.
- **Resumen del grupo** (⋮ → Resumen del grupo): alumnos, asistencia promedio, promedio general, faltas y alumnos en riesgo,
  con semáforo por alumno (rojo: calificación < 60 % o asistencia < 80 %; amarillo: < 80 % o < 90 %; verde: lo demás).
  *Vista previa* muestra el Excel tal cual; *Exportar Excel* genera el archivo con los indicadores y cada alumno coloreado, para el director.
- **Resumen por alumno** (⋮ → Resumen por alumno): se busca al alumno y se ve su calificación, cada apartado
  y el detalle día por día. *Exportar Excel* genera un archivo solo de ese alumno (para quejas o padres de familia).
- **Cuadernillo del grupo** (⋮ → Cuadernillo del grupo): el maestro sube un PDF por grupo y lo lee en el celular
  (página, zoom, botón "Ver grande" para ver solo el PDF en pantalla completa, recuerda dónde se quedó). Se guarda en el dispositivo (IndexedDB), funciona sin internet y
  se conserva aunque se vuelva a subir el Excel. Visor: PDF.js 3.11.174 en `vendor/pdfjs/`.
- **Confirmación de cambios**: si un alumno ya tenía marca y se elige otra, se pide confirmar.
- **Vista de lista** para revisar y corregir; tocar un nombre regresa al pase uno por uno desde ahí.
- **Agregar / quitar alumnos**: los nuevos se agregan al final de la lista del grupo.
  Desde ⋮ → Agregar alumno al grupo y ⋮ → Quitar alumno agregado; solo se pueden quitar los agregados desde la app (con buscador y confirmación).
- **Vista previa del Excel** dentro de la app, igual a lo que se exporta.
- **Exportar Excel**: copia del archivo original; por fecha van juntas Asistencia y Lectura, y Trabajos va aparte
  en un bloque a la derecha (separado por una columna vacía), todo
  a partir de la primera columna libre después del nombre.
  Valores por defecto: A = 1, F = 0, R = 2, J = 3; Lectura Sí = 1, No = 0; Trabajo 0/50/100
  (configurables en ⋮ → Valores de cada estado).
- Si se vuelve a subir un Excel exportado por la app, las fechas registradas se recuperan.

Los datos se guardan en el dispositivo (localStorage); el Excel solo se genera al exportar.
Funciona sin internet una vez abierta.

## Correr en local
```bash
python3 -m http.server 5173
```
y abrir http://localhost:5173

## Estructura
- `index.html`, `css/styles.css`, `js/app.js` — interfaz
- `js/excel-core.js` — lectura/escritura del Excel (ExcelJS 4.4.0 en `vendor/`)
- `manifest.json`, `sw.js`, `icons/` — instalación como app
