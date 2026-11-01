import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { buildDashboardModel, buildPivot, businessDeadline } from '../src/analytics.js';
import { parseAmount, parseDate, parseWorkbook } from '../src/data.js';

const utc = (date) => new Date(`${date}T00:00:00.000Z`);

const client = (key, contributor, periodStart = '2025-01-01') => ({
  key,
  contributor,
  tax: '0400722- Impuesto de prueba',
  taxKey: '0400722',
  frequency: 'Mensual',
  periodStart: utc(periodStart),
});

const payment = (key, contributor, amount, payDate, period = '2026-01-01') => ({
  key,
  contributor,
  tax: '0400722- Impuesto de prueba',
  taxKey: '0400722',
  amount,
  payDate: utc(payDate),
  period: utc(period),
  periodKey: period.slice(0, 7),
  dateBasis: 'period',
  dimensions: { tributo: '0400722- Impuesto de prueba' },
});

test('lee fechas y números con formatos de Excel y español', () => {
  assert.equal(parseDate(46120).toISOString(), '2026-04-08T00:00:00.000Z');
  assert.equal(parseDate('46120').toISOString(), '2026-04-08T00:00:00.000Z');
  assert.equal(parseDate('20260415').toISOString(), '2026-04-15T00:00:00.000Z');
  assert.equal(parseDate('15/04/2026').toISOString(), '2026-04-15T00:00:00.000Z');
  assert.equal(parseDate('2026-04').toISOString(), '2026-04-01T00:00:00.000Z');
  assert.equal(parseAmount('1.234,56'), 1234.56);
  assert.equal(parseAmount('$ 1,234.56'), 1234.56);
});

test('calcula el décimo día hábil y permite excluir feriados', () => {
  assert.equal(businessDeadline('2026-03', 10).toISOString(), '2026-03-13T00:00:00.000Z');
  assert.equal(businessDeadline('2026-03', 10, ['2026-03-02']).toISOString(), '2026-03-16T00:00:00.000Z');
});

test('distingue pago puntual, pago tardío y ausencia de pago', () => {
  const clients = [client('A', '10000000100- Alfa S.R.L.'), client('B', '10000000200- Beta S.R.L.'), client('C', '10000000300- Gamma S.R.L.')];
  const payments = [
    payment('A', '10000000100- Alfa S.R.L.', 100, '2026-01-14'),
    payment('A', '10000000100- Alfa S.R.L.', 50, '2026-01-15'),
    payment('B', '10000000200- Beta S.R.L.', 20, '2026-01-13'),
  ];
  const model = buildDashboardModel({
    clients,
    payments,
    period: '2026-01',
    taxFilter: '0400722',
    workingDays: 10,
    today: utc('2026-02-01'),
  });
  const states = new Map(model.rosterRows.map((row) => [row.key, row.status]));
  assert.equal(model.expectedCount, 3);
  assert.equal(model.lateCount, 1);
  assert.equal(model.statusCounts.paid, 1);
  assert.equal(model.unpaidCount, 1);
  assert.equal(states.get('A'), 'late');
  assert.equal(states.get('B'), 'paid');
  assert.equal(states.get('C'), 'unpaid');
  assert.equal(model.lateAmount, 50);
});

test('agrupa pagos dinámicamente por período y por contribuyente', () => {
  const clients = [client('A', '10000000100- Alfa S.R.L.'), client('B', '10000000200- Beta S.R.L.')];
  const payments = [
    payment('A', '10000000100- Alfa S.R.L.', 100, '2026-01-14'),
    payment('A', '10000000100- Alfa S.R.L.', 50, '2026-01-15'),
    payment('B', '10000000200- Beta S.R.L.', 20, '2026-01-13'),
  ];
  const byPeriod = buildPivot({ payments, clients, taxFilter: '0400722', dimension: 'period', metric: 'amount', rosterOnly: true });
  const byClient = buildPivot({ payments, clients, taxFilter: '0400722', dimension: 'contributor', metric: 'payments', scope: 'selected', period: '2026-01', rosterOnly: true });
  assert.deepEqual(byPeriod.values, [170]);
  assert.equal(byClient.rows.length, 2);
  assert.equal(byClient.rows[0].count, 2);
});

test('carga los dos Excel de muestra incluidos en el proyecto', async () => {
  const [clients, payments] = await Promise.all([
    parseWorkbook(await readFile('.xlsx/Cliente - obligación (vector fiscal).xlsx'), 'clients'),
    parseWorkbook(await readFile('.xlsx/Pago1 (persona jurídica).xlsx'), 'payments'),
  ]);
  assert.equal(clients.clients.length, 190);
  assert.equal(payments.payments.length, 17291);
  assert.ok(payments.dimensions.some((dimension) => dimension.key === 'dpa'));
  assert.ok(payments.payments.some((row) => row.periodKey === '2026-03'));
});
