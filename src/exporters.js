import { normalizeHeader } from './data.js';
import { formatMoney } from './formatting.js';

const PAPER_SIZES = {
  a4: { label: 'A4', excel: 9, pdf: 'a4' },
  letter: { label: 'Carta', excel: 1, pdf: 'letter' },
  a3: { label: 'A3', excel: 8, pdf: 'a3' },
  legal: { label: 'Legal', excel: 5, pdf: 'legal' },
  a5: { label: 'A5', excel: 11, pdf: 'a5' },
};

const HEADER_FILL = '173A50';
const HEADER_TEXT = 'FFFFFF';

function isMoneyHeader(value) {
  const normalized = normalizeHeader(value);
  return ['importe', 'cobrado', 'monto', 'valortributo', 'saldopendiente', 'deuda', 'recaudado', 'pagado', 'baseimponible']
    .some((hint) => normalized.includes(hint));
}

function printable(value) {
  if (value instanceof Date) {
    return new Intl.DateTimeFormat('es-CU', { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' }).format(value);
  }
  return String(value ?? '');
}

function fittedWidth(header, rows, columnIndex) {
  const widest = Math.max(
    Array.from(String(header ?? '')).length,
    ...rows.map((row) => Array.from(printable(row[columnIndex])).length),
  );
  return Math.max(12, Math.min(58, widest + 2));
}

function fileSafeName(value = 'seguimiento') {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9_-]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase() || 'seguimiento';
}

export function createFilename(title, period, extension) {
  const periodPart = period ? `-${period}` : '';
  return `${fileSafeName(title)}${periodPart}.${extension}`;
}

export async function createXlsxBlob({
  headers,
  rows,
  paperSize = 'a4',
  orientation = 'landscape',
  title = 'Seguimiento de pagos',
  subtitle = '',
}) {
  const ExcelJSModule = await import('exceljs');
  const ExcelJS = ExcelJSModule.default || ExcelJSModule;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ONAT · Disciplina de pago';
  workbook.lastModifiedBy = 'ONAT · Disciplina de pago';
  workbook.subject = title;
  workbook.title = title;
  workbook.created = new Date();
  workbook.modified = new Date();

  const sheet = workbook.addWorksheet('Seguimiento', {
    views: [{ state: 'frozen', ySplit: 1 }],
    pageSetup: {
      paperSize: PAPER_SIZES[paperSize]?.excel ?? PAPER_SIZES.a4.excel,
      orientation,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.25, right: 0.25, top: 0.55, bottom: 0.55, header: 0.2, footer: 0.25 },
    },
  });
  sheet.showGridLines = false;
  sheet.properties.defaultRowHeight = 19;
  sheet.columns = headers.map((header, index) => ({
    header,
    key: `column${index + 1}`,
    width: fittedWidth(header, rows, index),
  }));
  if (rows.length) sheet.addRows(rows);

  const header = sheet.getRow(1);
  header.height = 28;
  header.eachCell((cell) => {
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: HEADER_TEXT } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    cell.border = { bottom: { style: 'medium', color: { argb: '42A99D' } } };
  });

  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    row.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 9, ...(cell.font || {}) };
      cell.alignment = { vertical: 'middle', ...(cell.alignment || {}) };
      if (rowIndex % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F4F8FA' } };
    });
    const firstColumnCell = row.getCell(1);
    firstColumnCell.font = { name: 'Arial', size: 9, ...(firstColumnCell.font || {}), bold: true };
  }
  sheet.getColumn(1).eachCell((cell) => {
    cell.font = { ...(cell.font || {}), bold: true };
  });

  headers.forEach((columnHeader, index) => {
    const normalized = normalizeHeader(columnHeader);
    const column = sheet.getColumn(index + 1);
    if (isMoneyHeader(columnHeader)) column.numFmt = '[$$-409] #,##0.00';
    if (rows.some((row) => row[index] instanceof Date)) {
      column.numFmt = normalized.includes('periodo') ? 'mmm yyyy' : 'dd/mm/yyyy';
    }
    if (normalized.includes('cantidaddepagos')) column.numFmt = '0';
    if (normalized.includes('diasdeatraso')) column.numFmt = '0 "días"';
  });

  sheet.autoFilter = { from: 'A1', to: `${sheet.getColumn(sheet.columnCount).letter}${Math.max(1, sheet.rowCount)}` };
  sheet.pageSetup.printTitlesRow = '1:1';
  if (subtitle) sheet.headerFooter.oddFooter = `&L${subtitle}&R&P de &N`;
  if (title) sheet.headerFooter.oddHeader = `&L&"Arial,Bold"${title}`;

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export async function createPdfBlob({
  headers,
  rows,
  paperSize = 'a4',
  orientation = 'landscape',
  title = 'Seguimiento de pagos',
  subtitle = '',
  chartImage = null,
}) {
  const [jspdfModule, autoTableModule] = await Promise.all([
    import('jspdf'),
    import('jspdf-autotable'),
  ]);
  const { jsPDF } = jspdfModule;
  const autoTable = autoTableModule.default || autoTableModule.autoTable;
  const doc = new jsPDF({
    orientation,
    unit: 'mm',
    format: PAPER_SIZES[paperSize]?.pdf ?? PAPER_SIZES.a4.pdf,
    compress: true,
    putOnlyUsedFonts: true,
  });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = orientation === 'landscape' ? 9 : 12;

  doc.setTextColor(26, 48, 64);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.text(title, margin, 15);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(105, 123, 137);
  doc.setFontSize(8.5);
  const subtitleLines = subtitle ? doc.splitTextToSize(subtitle, pageWidth - margin * 2) : [];
  if (subtitleLines.length) doc.text(subtitleLines, margin, 21);
  doc.text(`Generado: ${new Intl.DateTimeFormat('es-CU').format(new Date())}`, pageWidth - margin, 15, { align: 'right' });
  let tableStartY = Math.max(28, 24 + Math.max(0, subtitleLines.length - 1) * 3.8 + 4);
  if (chartImage) {
    const imageProperties = doc.getImageProperties(chartImage);
    const maxChartWidth = pageWidth - margin * 2;
    const maxChartHeight = Math.max(40, Math.min(115, pageHeight - tableStartY - margin - 30));
    const imageScale = Math.min(maxChartWidth / imageProperties.width, maxChartHeight / imageProperties.height);
    const imageWidth = imageProperties.width * imageScale;
    const imageHeight = imageProperties.height * imageScale;
    const imageX = margin + (maxChartWidth - imageWidth) / 2;
    doc.addImage(chartImage, 'PNG', imageX, tableStartY, imageWidth, imageHeight, undefined, 'FAST');
    tableStartY += imageHeight + 7;
  }

  const amountColumns = new Set(headers.flatMap((header, index) => (isMoneyHeader(header) ? [index] : [])));
  const body = rows.map((row) => row.map((value, index) => (
    amountColumns.has(index) && typeof value === 'number' ? formatMoney(value) : printable(value)
  )));
  autoTable(doc, {
    head: [headers],
    body,
    startY: tableStartY,
    margin: { top: 24, right: margin, bottom: 13, left: margin },
    tableWidth: 'auto',
    theme: 'grid',
    showHead: 'everyPage',
    styles: {
      font: 'helvetica',
      fontSize: orientation === 'landscape' ? 7.4 : 6.6,
      cellPadding: 2,
      overflow: 'linebreak',
      valign: 'middle',
      lineColor: [221, 229, 234],
      lineWidth: 0.15,
      textColor: [55, 73, 88],
    },
    headStyles: {
      fillColor: [23, 58, 80],
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      fontSize: 7.2,
      cellPadding: 2.5,
    },
    alternateRowStyles: { fillColor: [244, 248, 250] },
    columnStyles: { 0: { fontStyle: 'bold' } },
    didDrawPage: (data) => {
      const pageNumber = doc.internal.getCurrentPageInfo().pageNumber;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(133, 148, 158);
      doc.text('ONAT · Disciplina de pago', margin, pageHeight - 6);
      doc.text(`Página ${pageNumber}`, pageWidth - margin, pageHeight - 6, { align: 'right' });
      if (data.pageNumber > 1 && title) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.setTextColor(26, 48, 64);
        doc.text(title, margin, 14);
      }
    },
  });

  return doc.output('blob');
}
