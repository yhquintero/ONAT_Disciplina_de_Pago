# ONAT · Disciplina de pago

Panel web local para revisar pagos de contribuyentes a partir de archivos Excel, identificar quién no tiene un pago registrado y explorar los cobros de **Importe ($)** con gráficos dinámicos.

## Iniciar el panel

Requiere Node.js 20.19+ o 22.12+.

```bash
npm install
npm run dev
```

Vite mostrará la dirección local (por defecto `http://localhost:5173`). Para validar la versión de producción:

```bash
npm test
npm run build
npm run preview
```

## Uso

El panel se abre con dos libros de ejemplo incluidos en `.xlsx/` y deja reemplazarlos desde **Importar Excel**:

1. **Padrón de clientes**: una hoja con la columna `Contribuyente` y todos los clientes que se deben revisar.
2. **Registro de pagos**: una hoja con `Contribuyente`, `Importe ($)` y `Fecha pago` / `Fecha de pago`.

Se admiten archivos `.xls`, `.xlsx`, `.xlsm` y `.xlsb`. El cruce prioriza el código de contribuyente que aparece antes del guion; si no está, usa el nombre normalizado. Cada libro puede tener varias hojas: se selecciona automáticamente la hoja que mejor coincide con los encabezados esperados.

### Regla de puntualidad

El valor inicial es **los primeros 10 días hábiles de cada mes**. En **Regla** se puede cambiar el día hábil límite y agregar los feriados que correspondan. Sábados y domingos se excluyen automáticamente; no se descarga ni se presupone un calendario oficial de feriados. Un pago en el día límite cuenta como puntual; uno posterior se clasifica como **Fuera de fecha**.

Cuando el archivo de pagos incluye `Período`, se usa el mes de esa obligación para calcular el vencimiento. Si no existe, se usa el mes de `Fecha de pago` y el panel avisa de esa suposición. Para pagos fraccionados, el cliente queda como **Fuera de fecha** si alguna de sus transacciones del período supera el vencimiento.

**“Sin pago”** significa que no aparece una transacción asociada en el registro importado. Se necesita el padrón completo para encontrar clientes sin pagos; el panel no inventa importes adeudados ni calcula saldos, porque eso requiere deuda esperada. Los resultados de plazo también deben revisarse frente a la periodicidad indicada en el padrón.

### Gráficos y exportación

- Resumen de importe cobrado, contribuyentes con pago, pagos fuera de fecha y clientes sin pago.
- Evolución por período, distribución puntual/tarde y ranking de contribuyentes.
- Explorador de gráficos: agrupa por período, fecha de pago, contribuyente, tributo y otras columnas detectadas; muestra suma de importe, número de pagos o contribuyentes únicos.
- Filtros de período y tributo; búsqueda por contribuyente; descarga de la vista o tabla en CSV.

## Privacidad

Los Excel elegidos por el usuario se leen en el navegador y no se suben a un backend. Los datos importados viven en memoria hasta que se recarga o cierra la pestaña; la regla y la lista de feriados sí se guardan en el almacenamiento local del navegador.

## Nota sobre los libros de ejemplo

El padrón de ejemplo describe el tributo como **trimestral**, mientras que la regla solicitada para el panel es mensual. Por eso se muestra una advertencia: revisa la periodicidad y el período seleccionado antes de usar el seguimiento para una actuación oficial.
