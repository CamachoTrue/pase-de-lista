# Pase de Lista

App web instalable (PWA) para pasar lista desde un Excel. Funciona en computadora y en celular.

## Funciones
- **Subir Excel**: cada hoja es un grupo; la columna **NOMBRE** se detecta sola.
- **Pase de lista uno por uno**: un alumno en pantalla con 4 opciones —
  **A** Asistencia, **F** Falta, **R** Retardo, **J** Justificado. Al elegir, pasa al siguiente.
  Empieza en el primer alumno sin marcar; botones Anterior / Saltar.
- **Confirmación de cambios**: si un alumno ya tenía marca y se elige otra, se pide confirmar.
- **Vista de lista** para revisar y corregir; tocar un nombre regresa al pase uno por uno desde ahí.
- **Agregar / quitar alumnos**: los nuevos se agregan al final de la lista del grupo.
  Solo se pueden quitar los agregados desde la app (con buscador y confirmación).
- **Vista previa del Excel** dentro de la app, igual a lo que se exporta.
- **Exportar Excel**: copia del archivo original con una columna por fecha
  (a partir de la primera columna libre después del nombre).
  Valores por defecto: A = 1, F = 0, R = 2, J = 3 (configurables en ⋮ → Valores de cada estado).
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
