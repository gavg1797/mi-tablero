# Mi Tablero de Estudio

App tipo Trello para organizar tus tareas de estudio. Es gratis, no tiene anuncios y funciona sin internet.

## Qué hace
- **Columnas:** Pendientes → Esta semana → En progreso → Finalizado (puedes renombrarlas, agregar más o cambiarlas de orden).
- **Arrastrar tarjetas:** con el mouse, o en el celular **manteniendo presionada** la tarjeta y moviéndola.
- **Check redondo:** márcalo y la tarea pasa sola a "Finalizado". Desmárcalo y regresa a su columna.
- **Checklist** dentro de cada tarea, con barra de progreso.
- **Materias con colores**, buscador y filtro por materia.
- **Vista Agenda:** tus tareas agrupadas en Vencidas, Hoy, Mañana, Próximos 7 días y Más adelante.
- **Paso automático a "Esta semana"** de las tareas que vencen en 7 días o menos.
- **Importar cronograma:** pega tu cronograma (una línea por tarea) y se crean todas las tarjetas.
- **Google Calendar:**
  - Botón **Google Calendar** en cada tarea: abre el evento ya listo, solo tocas *Guardar*.
  - Botón **.ics**, o *Opciones → Exportar a calendario*: crea un archivo con todas tus tareas y sus alarmas para importarlas de una vez.
- **Notificaciones** en el dispositivo y **funcionamiento sin internet**.
- **Respaldo:** descarga o restaura tus datos (sirve para pasarlos a otro dispositivo).

## Cómo publicarla gratis (una sola vez, unos 10 minutos)

### Opción A: GitHub Pages (recomendada, gratis para siempre)
1. Crea una cuenta gratis en https://github.com
2. Toca **New repository**, ponle de nombre `mi-tablero`, elige **Public** y toca **Create**.
3. Toca **uploading an existing file** y arrastra **todo el contenido** de esta carpeta (index.html, sw.js, manifest.webmanifest y los íconos .png). Luego toca **Commit changes**.
4. Ve a **Settings → Pages**. En *Branch* elige **main** y **/(root)**, y toca **Save**.
5. En 1 o 2 minutos tu app estará en `https://TU-USUARIO.github.io/mi-tablero/`

### Opción B: Netlify (también gratis)
1. Crea una cuenta gratis en https://app.netlify.com
2. Ve a **Add new site → Deploy manually** y arrastra esta carpeta.
3. Te da un enlace del tipo `https://algo.netlify.app` (puedes cambiarle el nombre).

> Ninguna de las dos cobra por un uso personal como este. No necesitas tarjeta de crédito.

## Instalarla en el celular
- **iPhone (Safari):** abre el enlace → botón Compartir ⬆︎ → **Agregar a pantalla de inicio**.
- **Android (Chrome):** abre el enlace → menú ⋮ → **Instalar app**.
- Después de instalarla, abre la app desde su ícono → ⋮ **Opciones → Activar notificaciones**.

## Sincronización con Google (ya configurada)
- App publicada en: https://gavg1797.github.io/mi-tablero/
- Toca la **nube ☁️** arriba → **Conectar con Google** → elige tu cuenta. Como la app está en modo "Prueba", Google mostrará el aviso *"Google no verificó esta app"*: toca **Continuar** (es tu propia app).
- Tus tareas se guardan en una carpeta oculta y privada de tu Google Drive y se combinan entre todos tus dispositivos. Sin internet sigues trabajando; al volver la conexión se sincroniza solo.
- Si la nube se pone naranja, la sesión de Google expiró (pasa más o menos cada hora): tócala para reconectar.
- Proyecto de Google Cloud: **Mi Tablero** (`mi-tablero-515000`). Solo las cuentas agregadas como *usuarios de prueba* (Google Auth Platform → Público) pueden conectarse.

## Importante: recordatorios y datos
- **Recordatorios garantizados:** usa Google Calendar. Sus alertas llegan aunque la app esté cerrada, sin internet y en todos tus dispositivos.
- **Notificaciones de la app:** llegan cuando la app está abierta o en segundo plano. Los teléfonos no permiten que una app web avise estando totalmente cerrada sin un servidor de pago; por eso los recordatorios fuertes van a Google Calendar.
- **Tus datos** se guardan en el dispositivo y, si conectas Google, también en tu Drive (gratis). Igual puedes descargar un respaldo en *Opciones*.
- Si borras los datos de navegación del navegador, se borran las tareas. Por eso conviene tener un respaldo.

## Varios tableros, columnas y zoom
- Toca **☰** (o el nombre del tablero) para abrir la **barra de menú**: ahí cambias de tablero, creas uno nuevo (se abre en otra pestaña), lo renombras, lo copias, lo eliminas y sincronizas uno o todos los tableros con Google.
- Cada tablero tiene su propia dirección (`…/mi-tablero/#b-xxxx`), así que puedes guardarlo como favorito o abrir varios a la vez.
- En cada columna, el botón **⋯** permite renombrar, ordenar por fecha, moverla, usarla como “Finalizado”, **ocultarla** o **eliminarla** (eliges a dónde mover sus tareas). Al final del tablero está **+ Añadir columna**.
- **Zoom** con los botones − / + (o Ctrl + rueda del mouse / Ctrl + y Ctrl −). Toca el porcentaje para volver al 100 %.

## Cargar tu cronograma (lectura inteligente)
Toca **⬇ Cargar cronograma** (o arrastra el archivo a la app) y elige tu archivo:
- **Excel** (.xlsx, .xls, .ods) y **CSV**: reconoce las columnas por su encabezado (Fecha, Curso/Materia, Actividad/Tema, Hora, Estado…). Si no hay encabezados, las deduce. Respeta las celdas combinadas y usa el nombre de cada hoja como materia.
- **Word** (.docx): lee tablas, listas y párrafos. Los títulos se usan como materia.
- **PDF** con texto (no escaneado): reconstruye renglones y columnas.
- **Texto** pegado o .txt.

Entiende fechas como `20/11`, `20/11/2026`, `2026-11-20`, `20 de noviembre`, `vie 21 nov`, `Nov 20`, rangos (`17 al 21 de noviembre`) y horas (`14:00`, `2 pm`, `3:00 p.m.`). Si una fila dice *Entregado/Hecho/✔* va a **Finalizado**; si dice *En progreso*, va a **En progreso**.

**Ubicación automática según la fecha de hoy:**
| Fecha | Columna |
|---|---|
| Hoy o ya vencida | En progreso |
| Próximos 7 días | Esta semana |
| Más adelante | Pendientes |
| Marcada como hecha (o fecha pasada, si eliges “marcarlas como finalizadas”) | Finalizado |

Cada día, al abrir la app, las tareas avanzan solas de columna a medida que se acercan sus fechas. Antes de importar ves una vista previa donde puedes corregir el nombre, la fecha, la materia o la columna. Las tareas que ya tienes no se duplican. Los recordatorios inteligentes avisan 2 días antes de los exámenes y 1 día antes del resto.

## Formato del cronograma
```
20/11 | Examen parcial | Cálculo
22/11/2026 14:00 | Entregar ensayo | Español
2026-11-25 - Leer capítulo 4 - Historia
```
Fecha (dd/mm, dd/mm/aaaa o aaaa-mm-dd), hora opcional (14:00 o 2:00 pm), tarea y materia.
