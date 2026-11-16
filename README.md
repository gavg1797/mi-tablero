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

## Importante: recordatorios y datos
- **Recordatorios garantizados:** usa Google Calendar. Sus alertas llegan aunque la app esté cerrada, sin internet y en todos tus dispositivos.
- **Notificaciones de la app:** llegan cuando la app está abierta o en segundo plano. Los teléfonos no permiten que una app web avise estando totalmente cerrada sin un servidor de pago; por eso los recordatorios fuertes van a Google Calendar.
- **Tus datos se guardan en cada dispositivo** (no en la nube, por eso no cuesta nada). Para pasarlos a otro celular: *Opciones → Descargar respaldo* y en el otro, *Restaurar respaldo*. Haz un respaldo de vez en cuando.
- Si borras los datos de navegación del navegador, se borran las tareas. Por eso conviene tener un respaldo.

## Formato del cronograma
```
20/11 | Examen parcial | Cálculo
22/11/2026 14:00 | Entregar ensayo | Español
2026-11-25 - Leer capítulo 4 - Historia
```
Fecha (dd/mm, dd/mm/aaaa o aaaa-mm-dd), hora opcional (14:00 o 2:00 pm), tarea y materia.
