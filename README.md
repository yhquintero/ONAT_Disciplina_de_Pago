# ONAT · Disciplina de pago

Panel web para cruzar un Patrón de contribuyentes con un registro de pagos, detectar pagos sin registrar y explorar los cobros de **Importe ($)** con gráficos dinámicos. La aplicación no tiene backend: los archivos que cada usuario selecciona se procesan dentro de su navegador.

## Empezar en local

Requiere Node.js 20.19.x o 22.12+ (se recomienda Node 22).

```bash
npm ci
npm run dev
```

El comando principal inicia HTTPS en **`https://localhost:5173`** con un certificado autofirmado de desarrollo; el navegador puede pedirte que aceptes la excepción local. Los libros de muestra de `.xlsx/` se cargan solo en desarrollo para facilitar las pruebas. No expongas el servidor de desarrollo a Internet.

Comandos alternativos:

```bash
npm run dev:http                 # HTTP local en http://localhost:5173
npm run dev:preview -- --port 4173  # HTTP para el preview de Arena
```

El preview de Arena debe usar el comando HTTP; `npm run dev` queda reservado para el acceso HTTPS directo desde tu equipo.

### Panel de Control PowerShell (Windows)

El repositorio incluye `scripts/PanelControl.ps1`, un menú local para iniciar la Web segura o la vista previa, instalar dependencias, ejecutar pruebas, compilar, revisar seguridad y consultar la hoja de ruta. No requiere permisos de administrador y no modifica los Excel importados ni publica cambios.

Desde la carpeta raíz del repositorio, abre PowerShell y ejecuta:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\PanelControl.ps1
```

También puedes usar `pwsh` si tienes PowerShell 7. El servidor que inicie el menú queda en otra ventana de terminal; ciérrala para detenerlo. Para HTTPS local, acepta la excepción del certificado autofirmado si tu navegador la muestra.

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

1. **Patrón de clientes**: una hoja con la columna `Contribuyente` y todos los clientes que se deben revisar.
2. **Registro de pagos**: una hoja con `Contribuyente`, `Importe ($)` y `Fecha pago` / `Fecha de pago`.

Se admiten archivos `.xls`, `.xlsx`, `.xlsm` y `.xlsb`. Si hay varias hojas, se selecciona automáticamente la que mejor coincide con los encabezados. El cruce prioriza el código de contribuyente que aparece antes del guion; si falta, usa el nombre normalizado. Cuando el registro contiene la columna `Período`, se usa ese mes como período de obligación. Si no está, se usa el mes de `Fecha de pago` y la interfaz lo advierte.

## Regla de puntualidad

El valor inicial es **los primeros 10 días hábiles de cada mes**. En **Regla** se puede cambiar el día hábil límite y agregar feriados. Sábados y domingos se excluyen automáticamente; no se descarga ni se presume un calendario oficial. Un pago en el día límite cuenta como puntual; uno posterior se clasifica como **Fuera de fecha**. Para pagos fraccionados, el cliente se marca fuera de fecha si alguna transacción del período supera el vencimiento.

**“Sin pago”** significa que no aparece una transacción coincidente en el registro importado. Se requiere el Patrón completo para identificar esos casos. El panel no inventa importes adeudados ni calcula saldos porque eso requiere conocer la deuda esperada. Revisa también la periodicidad definida en el Patrón antes de usar los estados para una actuación oficial.

## Funcionalidad

- Indicadores de importe recaudado, contribuyentes con pago, pagos tardíos y clientes sin pago registrado.
- Gráficos de evolución por período, puntualidad y ranking de contribuyentes.
- Explorador dinámico: agrupa por período, fecha de pago, contribuyente, tributo y otras columnas detectadas; muestra suma de importe, número de pagos o contribuyentes únicos. Permite gráficos de barras, líneas, área y dona, además de limitar el número de grupos visibles.
- Los importes se presentan con el símbolo al inicio y el formato solicitado, por ejemplo **`$ 1,000,000.00`**. Los importes de Excel y PDF siguen la misma convención.
- La importación conserva los campos originales de ambos libros, incluidos atributos adicionales como DPA, forma organizativa, organismo de dirección, configuración/periodicidad, valor del tributo, período inicial y fecha de alta. El modal muestra las columnas detectadas; las exportaciones de seguimiento incluyen atributos del Patrón y los movimientos pueden exportarse con sus columnas originales, período analizado y fila/hoja de origen. `Valor tributo` se conserva tal como viene en Excel, pero no se interpreta como deuda ni se suma al importe cobrado.
- Exportación a `.xlsx`, PDF y CSV. Excel aplica negrita a la primera columna, ajusta anchos automáticamente y configura la impresión; PDF y Excel admiten A4, Carta, A3, Legal y A5, en orientación horizontal o vertical. El PDF del explorador incluye el gráfico visible y su tabla. Los movimientos detallados se ofrecen en XLSX y CSV para evitar informes PDF excesivamente voluminosos.
- Filtros de período y tributo, búsqueda por contribuyente y preferencias de vencimiento/feriados guardadas localmente. Los libros importados viven en memoria hasta recargar o cerrar la pestaña.

## Ideas para mejoras futuras (por etapas)

Estas propuestas no se activan automáticamente; se pueden abordar una por una según prioridad:

1. **Validación de importación:** detectar duplicados, filas incompletas, identificadores ambiguos, fechas inválidas y períodos sospechosos; ofrecer una vista previa y un informe descargable antes de analizar.
2. **Mapeo de columnas:** guardar equivalencias configurables para encabezados distintos, seleccionar la hoja y decidir qué campos ignorar o incluir en cada libro.
3. **Conciliación de contribuyentes:** mostrar coincidencias exactas, posibles duplicados y pagos sin correspondencia con el Patrón antes de aplicar el cruce.
4. **Análisis de atributos del Patrón:** filtrar y comparar por configuración/periodicidad, valor del tributo, antigüedad y futuros campos extraídos.
5. **Comparación temporal:** comparar meses o años y mostrar variaciones de importe, puntualidad y contribuyentes pendientes.
6. **Trazabilidad del cálculo:** explicar por fila el vencimiento aplicado, los días de atraso y las transacciones que originaron el estado, usando la fila y hoja de origen ya conservadas.
7. **Revisión de casos y respaldo local:** añadir notas y estado de revisión por contribuyente, con exportación/importación de una copia de trabajo cifrada o descargable.
8. **Accesibilidad y rendimiento:** atajos de teclado, mejores ayudas de lectura, paginación de tablas grandes y métricas para medir la velocidad de importación.
9. **Uso compartido opcional:** si se necesita colaboración, evaluar perfiles, permisos y copias de respaldo; los datos fiscales requerirían un backend protegido y no deberían publicarse en GitHub Pages.
10. **Alertas locales:** recordar vencimientos próximos y atrasos según la regla y los feriados configurados, sin transmitir datos.
11. **Plantillas de informes:** guardar filtros, columnas y tamaños de página preferidos por tributo o periodicidad.
12. **Modo presentación:** mostrar indicadores agregados y ocultar nombres/códigos para reuniones o capacitación.

Como siguiente mejora recomendada, empezar por **validación y conciliación de importación**: revisar duplicados, campos vacíos y pagos sin correspondencia antes de que afecten al análisis, manteniendo el procesamiento local.

## Arquitectura

- `src/data.js`: lectura de Excel y normalización de encabezados, importada bajo demanda.
- `src/analytics.js`: cálculo de vencimientos, cruce y métricas.
- `src/exporters.js`: creación local de libros `.xlsx` y PDF con formato de impresión.
- `src/formatting.js`: formato monetario común para la interfaz y reportes.
- `src/App.jsx` / `src/styles.css`: interfaz, importación, exportaciones y gráficos.
- `scripts/PanelControl.ps1`: menú local de PowerShell para gestionar desarrollo, pruebas y compilación.
- `tests/core.test.js`, `tests/exporters.test.js` y `tests/formatting.test.js`: pruebas de lectura de libros, extracción de columnas, fechas, métricas, exportaciones y formato monetario.
- `.github/workflows/ci.yml` y `pages.yml`: verificación y publicación.
