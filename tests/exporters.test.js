import test from 'node:test';
import assert from 'node:assert/strict';
import { createFilename, createPdfBlob, createXlsxBlob } from '../src/exporters.js';

const headers = [
  'Contribuyente',
  'Tributo',
  'Estado',
  'Importe cobrado ($)',
  'Cantidad de pagos',
  'Primer pago',
  'Último pago',
  'Días de atraso',
];
const rows = [[
  '10000000100- Alfa S.R.L.',
  '0400722- Impuesto de prueba',
  'En fecha',
  1250.5,
  2,
  new Date('2026-03-10T00:00:00.000Z'),
  new Date('2026-03-11T00:00:00.000Z'),
  0,
]];

test('crea XLSX con los encabezados, primera columna en negrita y página ajustada', async () => {
  const blob = await createXlsxBlob({ headers, rows, paperSize: 'letter', orientation: 'landscape', title: 'Seguimiento' });
  assert.equal(blob.type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const ExcelJSModule = await import('exceljs');
  const ExcelJS = ExcelJSModule.default || ExcelJSModule;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await blob.arrayBuffer());
  const sheet = workbook.getWorksheet('Seguimiento');
  assert.deepEqual(sheet.getRow(1).values.slice(1), headers);
  assert.equal(sheet.getCell('A2').font.bold, true);
  assert.ok(sheet.getColumn(1).width >= '10000000100- Alfa S.R.L.'.length);
  assert.equal(sheet.pageSetup.paperSize, 1);
  assert.equal(sheet.pageSetup.orientation, 'landscape');
  assert.equal(sheet.pageSetup.fitToWidth, 1);
});

test('crea PDF con orientación y tamaño de página seleccionados', async () => {
  const blob = await createPdfBlob({ headers, rows, paperSize: 'a4', orientation: 'landscape', title: 'Seguimiento' });
  assert.equal(blob.type, 'application/pdf');
  assert.match(await blob.slice(0, 8).text(), /^%PDF-/);
  assert.ok(blob.size > 1000);
});

test('genera nombres de archivo seguros con el período', () => {
  assert.equal(createFilename('Seguimiento de pagos', '2026-03', 'xlsx'), 'seguimiento-de-pagos-2026-03.xlsx');
});
