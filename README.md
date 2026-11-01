# ONAT · Disciplina de pago

Panel web para cruzar un padrón de contribuyentes con un registro de pagos, detectar pagos sin registrar y explorar los cobros de **Importe ($)** con gráficos dinámicos. La aplicación no tiene backend: los archivos que cada usuario selecciona se procesan dentro de su navegador.

## Empezar en local

Requiere Node.js 20.19.x o 22.12+ (se recomienda Node 22).

```bash
npm ci
npm run dev
```

Vite mostrará una dirección local, normalmente `http://localhost:5173`. En desarrollo local se cargan los dos libros de muestra de `.xlsx/` para facilitar las pruebas. El modo de desarrollo no debe exponerse a Internet.

Comprobación antes de integrar cambios:

```bash
npm run check   # pruebas + compilación + comprobación del artefacto público
npm audit
```

## Publicar con HTTPS

El repositorio incluye un flujo de GitHub Actions en `.github/workflows/pages.yml`. Después de integrar el cambio en `main`:

1. En GitHub abre **Settings → Pages** y selecciona **GitHub Actions** como *Build and deployment source* si aún no está configurado.
2. El flujo de publicación ejecuta las pruebas y compila el sitio bajo la ruta del repositorio.
3. Cuando finalice correctamente, la URL del entorno `github-pages` será `https://yhquintero.github.io/ONAT_Disciplina_de_Pago/`. En **Settings → Pages**, activa **Enforce HTTPS** cuando GitHub lo ofrezca.

`VITE_BASE_PATH` permite adaptar la compilación a otra ruta o proveedor. Para probar localmente el build de Pages:

```bash
VITE_BASE_PATH=/ONAT_Disciplina_de_Pago/ npm run build
npm run preview
# abrir http://localhost:4173/ONAT_Disciplina_de_Pago/
```

El flujo `.github/workflows/ci.yml` también valida pruebas y compilación en cada push y pull request. La aplicación publicada comienza vacía de forma intencional: los libros de ejemplo solo se leen en desarrollo y no se empaquetan dentro del artefacto público. Los usuarios cargan sus propios libros desde **Importar Excel**.

> **HTTPS no significa acceso privado.** GitHub Pages es un sitio público y no incorpora inicio de sesión. Los libros que un usuario carga desde la página no se suben al sitio, pero cualquier archivo ya versionado en un repositorio público sí es accesible públicamente. Los Excel de muestra actuales están en este repositorio público; no agregues datos fiscales reales al repositorio. Si contienen información sensible, retíralos del historial público o vuelve privado el repositorio antes de compartir la URL. Para colaboración entre usuarios con permisos o datos centralizados haría falta agregar autenticación y un backend, no solo HTTPS.

## Formato de datos

Importa por separado:

1. **Padrón de clientes**: una hoja con la columna `Contribuyente` y todos los clientes que se deben revisar.
2. **Registro de pagos**: una hoja con `Contribuyente`, `Importe ($)` y `Fecha pago` / `Fecha de pago`.

Se admiten archivos `.xls`, `.xlsx`, `.xlsm` y `.xlsb`. Si hay varias hojas, se selecciona automáticamente la que mejor coincide con los encabezados. El cruce prioriza el código de contribuyente que aparece antes del guion; si falta, usa el nombre normalizado. Cuando el registro contiene la columna `Período`, se usa ese mes como período de obligación. Si no está, se usa el mes de `Fecha de pago` y la interfaz lo advierte.

## Regla de puntualidad

El valor inicial es **los primeros 10 días hábiles de cada mes**. En **Regla** se puede cambiar el día hábil límite y agregar feriados. Sábados y domingos se excluyen automáticamente; no se descarga ni se presume un calendario oficial. Un pago en el día límite cuenta como puntual; uno posterior se clasifica como **Fuera de fecha**. Para pagos fraccionados, el cliente se marca fuera de fecha si alguna transacción del período supera el vencimiento.

**“Sin pago”** significa que no aparece una transacción coincidente en el registro importado. Se requiere el padrón completo para identificar esos casos. El panel no inventa importes adeudados ni calcula saldos porque eso requiere conocer la deuda esperada. Revisa también la periodicidad definida en el padrón antes de usar los estados para una actuación oficial.

## Funcionalidad

- Indicadores de importe recaudado, contribuyentes con pago, pagos tardíos y clientes sin pago registrado.
- Gráficos de evolución por período, puntualidad y ranking de contribuyentes.
- Explorador dinámico: agrupa por período, fecha de pago, contribuyente, tributo y otras columnas detectadas; muestra suma de importe, número de pagos o contribuyentes únicos.
- Filtros de período y tributo, búsqueda por contribuyente y exportación CSV.
- Preferencias de vencimiento y feriados guardadas en el almacenamiento local del navegador. Los libros importados viven en memoria hasta recargar o cerrar la pestaña.

## Arquitectura

- `src/data.js`: lectura de Excel y normalización de encabezados, importada bajo demanda.
- `src/analytics.js`: cálculo de vencimientos, cruce y métricas.
- `src/App.jsx` / `src/styles.css`: interfaz y gráficos.
- `tests/core.test.js`: pruebas de fechas, clasificación, pivotes y libros de muestra.
- `.github/workflows/ci.yml` y `pages.yml`: verificación y publicación.
