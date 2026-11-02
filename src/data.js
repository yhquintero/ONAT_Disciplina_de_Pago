export const normalizeText = (value = '') =>
  String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();

export const normalizeHeader = (value = '') => normalizeText(value).replace(/\s+/g, '');

export function contributorKey(value = '') {
  const text = String(value ?? '').trim();
  const taxNumber = text.match(/^([a-z0-9]{6,})\s*[-–—]/i)?.[1]
    || text.match(/^(\d{6,})\b/)?.[1];
  return taxNumber ? taxNumber.toUpperCase() : normalizeText(text);
}

export function taxKey(value = '') {
  const text = String(value ?? '').trim();
  const code = text.match(/^([a-z0-9.]{3,})\s*[-–—]/i)?.[1]
    || text.match(/^(\d{4,})\b/)?.[1];
  return code ? code.replace(/[^a-z0-9]/gi, '').toUpperCase() : normalizeText(text);
}

export function monthKey(date) {
  if (!date) return '';
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}

export function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value > 100000000000) {
      const date = new Date(value);
      return Number.isNaN(date.getTime()) ? null : new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    }
    if (value >= 10000101 && value <= 99991231 && Number.isInteger(value)) {
      const text = String(value);
      if (text.length === 8) return validDate(Number(text.slice(0, 4)), Number(text.slice(4, 6)), Number(text.slice(6, 8)));
    }
    return excelSerialDate(value);
  }

  const text = String(value ?? '').trim();
  if (!text) return null;
  if (/^\d{8}$/.test(text)) {
    return validDate(Number(text.slice(0, 4)), Number(text.slice(4, 6)), Number(text.slice(6, 8)));
  }
  if (/^\d{1,6}(?:\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 0 && serial <= 300000) return excelSerialDate(serial);
  }

  if (/^\d{4}[-/]\d{1,2}$/.test(text)) {
    const [year, month] = text.split(/[-/]/).map(Number);
    return month >= 1 && month <= 12 ? new Date(Date.UTC(year, month - 1, 1)) : null;
  }

  const ymd = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (ymd) {
    const [, year, month, day] = ymd.map(Number);
    return validDate(year, month, day);
  }

  // ONAT exports commonly use dd/mm/yyyy; interpret slash and dot formats that way.
  const dmy = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$/);
  if (dmy) {
    let [, day, month, year] = dmy.map(Number);
    if (year < 100) year += year < 50 ? 2000 : 1900;
    return validDate(year, month, day);
  }

  const parsed = Date.parse(text);
  if (Number.isNaN(parsed)) return null;
  const date = new Date(parsed);
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function excelSerialDate(value) {
  // Excel's 1900 system has a fictitious 29-Feb-1900; the different epoch below
  // preserves modern serial dates and also handles the handful of early dates.
  if (value < 1 || value > 300000) return null;
  const wholeDays = Math.floor(value);
  const epoch = wholeDays < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
  return new Date(epoch + wholeDays * 86400000);
}

function validDate(year, month, day) {
  const result = new Date(Date.UTC(year, month - 1, day));
  if (result.getUTCFullYear() !== year || result.getUTCMonth() !== month - 1 || result.getUTCDate() !== day) return null;
  return result;
}

export function parseAmount(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const raw = String(value ?? '').trim();
  if (!raw) return null;

  let cleaned = raw
    .replace(/[\s\u00a0]/g, '')
    .replace(/\$/g, '')
    .replace(/[^\d,.-]/g, '');
  if (!cleaned || !/\d/.test(cleaned)) return null;

  const lastComma = cleaned.lastIndexOf(',');
  const lastDot = cleaned.lastIndexOf('.');
  if (lastComma >= 0 && lastDot >= 0) {
    const decimalMark = lastComma > lastDot ? ',' : '.';
    const thousandsMark = decimalMark === ',' ? '.' : ',';
    cleaned = cleaned.split(thousandsMark).join('');
    if (decimalMark === ',') cleaned = cleaned.replace(',', '.');
  } else if (lastComma >= 0) {
    const decimals = cleaned.length - lastComma - 1;
    cleaned = decimals > 0 && decimals <= 2
      ? cleaned.replace(',', '.')
      : cleaned.replace(/,/g, '');
  } else if ((cleaned.match(/\./g) || []).length > 1) {
    const finalDot = cleaned.lastIndexOf('.');
    const decimals = cleaned.length - finalDot - 1;
    cleaned = decimals > 0 && decimals <= 2
      ? `${cleaned.slice(0, finalDot).replace(/\./g, '')}.${cleaned.slice(finalDot + 1)}`
      : cleaned.replace(/\./g, '');
  }

  const amount = Number(cleaned);
  return Number.isFinite(amount) ? amount : null;
}

// String.prototype.replace is not an index fallback; keep these tiny helpers explicit.
function columnIndex(headers, kind) {
  const exactMatches = {
    contributor: ['contribuyente', 'cliente', 'nombrecliente', 'razonsocial'],
    period: ['periodo', 'periodofiscal', 'mesperiodo', 'mesdeobligacion', 'periododeliquidacion'],
    createdAt: ['fechadealta', 'fechaalta', 'alta'],
  };
  const includeMatches = {
    contributor: ['contribuyente', 'razonsocial', 'nombrecliente', 'cliente'],
    amount: ['importe', 'montopagado', 'monto', 'totalcobrado', 'pagado'],
    paymentDate: ['fechadepago', 'fechapago', 'fechacobro', 'fechadecobro', 'paymentdate'],
    period: ['periodofiscal', 'mesdeobligacion', 'periododeliquidacion', 'mesperiodo'],
    tax: ['tributo', 'impuesto', 'conceptotributario'],
    periodicity: ['periodicidad', 'configuracion', 'frecuencia'],
    periodStart: ['periodoinicial', 'iniciodelperiodo', 'periododeinicio'],
    createdAt: ['fechaalta', 'fechadealta'],
  };
  const normalized = headers.map(normalizeHeader);
  const exact = exactMatches[kind] ? normalized.findIndex((value) => exactMatches[kind].includes(value)) : -1;
  if (exact >= 0) return exact;
  return normalized.findIndex((value) => (includeMatches[kind] || []).some((candidate) => value.includes(candidate)));
}

function locateHeader(rows, role) {
  let best = null;
  const scanCount = Math.min(rows.length, 25);
  for (let index = 0; index < scanCount; index += 1) {
    const row = rows[index] || [];
    const headers = row.map((value) => String(value ?? '').trim());
    const columns = {
      contributor: columnIndex(headers, 'contributor'),
      amount: columnIndex(headers, 'amount'),
      paymentDate: columnIndex(headers, 'paymentDate'),
      period: columnIndex(headers, 'period'),
      tax: columnIndex(headers, 'tax'),
      periodicity: columnIndex(headers, 'periodicity'),
      periodStart: columnIndex(headers, 'periodStart'),
      createdAt: columnIndex(headers, 'createdAt'),
    };
    if (columns.contributor < 0) continue;

    const score = 4
      + (columns.amount >= 0 ? 2 : 0)
      + (columns.paymentDate >= 0 ? 3 : 0)
      + (columns.period >= 0 ? 1 : 0)
      + (columns.tax >= 0 ? 1 : 0);
    if (!best || score > best.score) best = { index, headers, columns, score };
  }

  if (!best) return null;
  if (role === 'payments' && (best.columns.amount < 0 || best.columns.paymentDate < 0)) return null;
  return best;
}

function getDimensionColumns(headers, columns) {
  const core = new Set([
    columns.contributor,
    columns.amount,
    columns.paymentDate,
    columns.period,
  ].filter((index) => index >= 0));
  const seen = new Set();
  return headers.reduce((items, label, index) => {
    const key = normalizeHeader(label) || `campo${index + 1}`;
    if (core.has(index) || !String(label ?? '').trim() || seen.has(key)) return items;
    seen.add(key);
    items.push({ key, label: String(label).trim(), index });
    return items;
  }, []);
}

function cellText(value) {
  return String(value ?? '').trim();
}

function getSourceFields(headers, dataRows) {
  const labelCounts = new Map();
  return headers.reduce((fields, header, index) => {
    const sourceLabel = cellText(header);
    const hasData = dataRows.some((row) => row?.[index] !== '' && row?.[index] != null);
    if (!sourceLabel && !hasData) return fields;

    const displayLabel = sourceLabel || `Columna ${index + 1} (sin encabezado)`;
    const normalized = normalizeHeader(sourceLabel) || `campo${index + 1}`;
    const occurrence = (labelCounts.get(normalized) || 0) + 1;
    labelCounts.set(normalized, occurrence);
    fields.push({
      key: occurrence === 1 ? normalized : `${normalized}_${occurrence}`,
      label: occurrence === 1 ? displayLabel : `${displayLabel} (${occurrence})`,
      sourceLabel,
      index,
    });
    return fields;
  }, []);
}

function normalizeSourceValue(value, label) {
  if (value == null || value === '') return '';
  const normalized = normalizeHeader(label);
  if (value instanceof Date) return parseDate(value) || value;
  if (/(fecha|date|periodo)/.test(normalized)) return parseDate(value) || cellText(value);
  return typeof value === 'string' ? value.trim() : value;
}

function sourceDimensionValue(value, label) {
  const normalized = normalizeSourceValue(value, label);
  if (normalized instanceof Date) {
    return `${normalized.getUTCFullYear()}-${String(normalized.getUTCMonth() + 1).padStart(2, '0')}-${String(normalized.getUTCDate()).padStart(2, '0')}`;
  }
  return cellText(normalized);
}

function extractSourceData(row, fields) {
  return Object.fromEntries(fields.flatMap(({ key, sourceLabel, label, index }) => {
    const value = normalizeSourceValue(row?.[index], sourceLabel || label);
    return value === '' ? [] : [[key, value]];
  }));
}

function parseClientRows(rows, headerInfo, fileName, sheetName) {
  const { index: headerIndex, columns, headers } = headerInfo;
  const dataRows = rows.slice(headerIndex + 1);
  const fields = getSourceFields(headers, dataRows);
  const clients = [];
  let skippedRows = 0;

  dataRows.forEach((row, rowOffset) => {
    const contributor = cellText(row[columns.contributor]);
    if (!contributor) return;
    const key = contributorKey(contributor);
    if (!key) {
      skippedRows += 1;
      return;
    }
    const tax = columns.tax >= 0 ? cellText(row[columns.tax]) : '';
    const periodStart = columns.periodStart >= 0 ? parseDate(row[columns.periodStart]) : null;
    const createdAt = columns.createdAt >= 0 ? parseDate(row[columns.createdAt]) : null;
    const frequency = columns.periodicity >= 0 ? cellText(row[columns.periodicity]) : '';
    clients.push({
      key,
      contributor,
      tax,
      taxKey: taxKey(tax),
      periodStart,
      createdAt,
      frequency,
      sourceData: extractSourceData(row, fields),
      sourceRow: headerIndex + rowOffset + 2,
      fileName,
      sheetName,
    });
  });

  if (!clients.length) throw new Error('No encontré filas con un valor en la columna Contribuyente.');
  return { clients, skippedRows, sourceRows: rows.length - headerIndex - 1, headers, fields };
}

function parsePaymentRows(rows, headerInfo, fileName, sheetName) {
  const { index: headerIndex, columns, headers } = headerInfo;
  const dataRows = rows.slice(headerIndex + 1);
  const fields = getSourceFields(headers, dataRows);
  const dimensions = getDimensionColumns(headers, columns);
  const payments = [];
  let skippedRows = 0;

  dataRows.forEach((row, rowOffset) => {
    const contributor = cellText(row[columns.contributor]);
    const amount = parseAmount(row[columns.amount]);
    const payDate = parseDate(row[columns.paymentDate]);
    if (!contributor && amount == null && !payDate) return;
    if (!contributor || amount == null || !payDate) {
      skippedRows += 1;
      return;
    }

    const period = columns.period >= 0 ? parseDate(row[columns.period]) : null;
    const tax = columns.tax >= 0 ? cellText(row[columns.tax]) : '';
    const dimensionValues = {};
    dimensions.forEach(({ key, label, index }) => {
      dimensionValues[key] = sourceDimensionValue(row[index], label);
    });

    payments.push({
      key: contributorKey(contributor),
      contributor,
      tax,
      taxKey: taxKey(tax),
      amount,
      payDate,
      period,
      periodKey: monthKey(period || payDate),
      dateBasis: period ? 'period' : 'payment',
      dimensions: dimensionValues,
      sourceData: extractSourceData(row, fields),
      sourceRow: headerIndex + rowOffset + 2,
      fileName,
      sheetName,
    });
  });

  if (!payments.length) throw new Error('No encontré pagos válidos con Contribuyente, Importe ($) y Fecha de pago.');
  return {
    payments,
    dimensions: dimensions.map(({ key, label }) => ({ key, label })),
    skippedRows,
    sourceRows: rows.length - headerIndex - 1,
    headers,
    fields,
  };
}

export async function parseWorkbook(buffer, role, fileName = 'Libro de Excel') {
  const XLSX = await import('@e965/xlsx');
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, dense: true });
  let bestMatch = null;
  let bestSheetName = '';

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '', blankrows: false });
    const headerInfo = locateHeader(rows, role);
    if (!headerInfo) continue;
    const score = headerInfo.score + Math.min(rows.length / 10000, 1);
    if (!bestMatch || score > bestMatch.score) {
      bestMatch = { rows, headerInfo, score };
      bestSheetName = sheetName;
    }
  }

  if (!bestMatch) {
    if (role === 'payments') {
      throw new Error('No identifiqué una hoja con Contribuyente, Importe ($) y Fecha de pago. Revisa los encabezados del Excel.');
    }
    throw new Error('No identifiqué una hoja con la columna Contribuyente. Revisa los encabezados del Excel.');
  }

  const result = role === 'payments'
    ? parsePaymentRows(bestMatch.rows, bestMatch.headerInfo, fileName, bestSheetName)
    : parseClientRows(bestMatch.rows, bestMatch.headerInfo, fileName, bestSheetName);

  return {
    ...result,
    sheetName: bestSheetName,
    fileName,
  };
}
