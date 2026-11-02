import { monthKey } from './data.js';

export function dateKey(date) {
  if (!date) return '';
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

export function dateFromKey(key) {
  if (!key) return null;
  const [year, month, day = '1'] = key.split('-').map(Number);
  if (!year || !month || !day) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

export function monthLabel(key, style = 'long') {
  const date = dateFromKey(`${key}-01`);
  if (!date) return key;
  const text = new Intl.DateTimeFormat('es-CU', { month: style, year: 'numeric', timeZone: 'UTC' }).format(date);
  return text.charAt(0).toUpperCase() + text.slice(1);
}

export function shortMonthLabel(key) {
  const date = dateFromKey(`${key}-01`);
  return date
    ? new Intl.DateTimeFormat('es-CU', { month: 'short', year: '2-digit', timeZone: 'UTC' }).format(date).replace('.', '')
    : key;
}

export function formatDate(date, options = {}) {
  if (!date) return '—';
  return new Intl.DateTimeFormat('es-CU', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    timeZone: 'UTC',
    ...options,
  }).format(date).replace('.', '');
}

export function businessDeadline(period, workingDays = 10, holidays = []) {
  if (!period) return null;
  const [year, month] = period.split('-').map(Number);
  if (!year || !month) return null;
  const target = Math.max(1, Math.min(31, Number(workingDays) || 10));
  const holidaySet = new Set(holidays);
  const day = new Date(Date.UTC(year, month - 1, 1));
  let count = 0;

  while (count < target && day.getUTCMonth() === month - 1) {
    const key = dateKey(day);
    const weekday = day.getUTCDay();
    if (weekday !== 0 && weekday !== 6 && !holidaySet.has(key)) count += 1;
    if (count < target) day.setUTCDate(day.getUTCDate() + 1);
  }
  return count === target ? day : null;
}

function isActiveInPeriod(client, period) {
  if (!client.periodStart) return true;
  return monthKey(client.periodStart) <= period;
}

function matchesTax(taxCode, filter) {
  return !filter || !taxCode || taxCode === filter;
}

function activeRosterMap(clients, period, taxFilter) {
  const map = new Map();
  clients.forEach((client) => {
    if (!client.key || !isActiveInPeriod(client, period)) return;
    if (taxFilter && client.taxKey && client.taxKey !== taxFilter) return;
    if (!map.has(client.key)) map.set(client.key, []);
    map.get(client.key).push(client);
  });
  return map;
}

function paymentMatchesRoster(payment, rosterEntries, taxFilter) {
  if (!rosterEntries) return false;
  if (!matchesTax(payment.taxKey, taxFilter)) return false;
  if (!payment.taxKey) return true;
  return rosterEntries.some((client) => !client.taxKey || client.taxKey === payment.taxKey);
}

export function getPaymentsForPeriod({ payments, clients, period, taxFilter = '', rosterOnly = true }) {
  if (!period) return [];
  const monthPayments = payments.filter((payment) => payment.periodKey === period && matchesTax(payment.taxKey, taxFilter));
  if (!rosterOnly || !clients.length) return monthPayments;
  const roster = activeRosterMap(clients, period, taxFilter);
  return monthPayments.filter((payment) => paymentMatchesRoster(payment, roster.get(payment.key), taxFilter));
}

export function buildDashboardModel({
  clients = [],
  payments = [],
  period = '',
  taxFilter = '',
  workingDays = 10,
  holidays = [],
  today = new Date(),
}) {
  const rosterLoaded = clients.length > 0;
  const deadline = businessDeadline(period, workingDays, holidays);
  const dueKey = deadline ? dateKey(deadline) : '';
  const todayDate = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
  const overdue = Boolean(deadline && todayDate.getTime() > deadline.getTime());
  const expectedEntries = clients.filter((client) => {
    if (!client.key || !period || !isActiveInPeriod(client, period)) return false;
    return !taxFilter || !client.taxKey || client.taxKey === taxFilter;
  });
  const rosterMap = new Map();
  expectedEntries.forEach((client) => {
    if (!rosterMap.has(client.key)) rosterMap.set(client.key, []);
    rosterMap.get(client.key).push(client);
  });

  const periodPayments = getPaymentsForPeriod({ payments, clients, period, taxFilter, rosterOnly: rosterLoaded });
  if (!rosterLoaded) {
    periodPayments.forEach((payment) => {
      if (!rosterMap.has(payment.key)) rosterMap.set(payment.key, []);
    });
  }

  const paymentMap = new Map();
  periodPayments.forEach((payment) => {
    if (!paymentMap.has(payment.key)) paymentMap.set(payment.key, []);
    paymentMap.get(payment.key).push(payment);
  });

  const rosterRows = Array.from(rosterMap.entries()).map(([key, entries]) => {
    const clientPayments = paymentMap.get(key) || [];
    const representative = entries[0] || {};
    const total = clientPayments.reduce((sum, payment) => sum + payment.amount, 0);
    const latePayments = deadline
      ? clientPayments.filter((payment) => payment.payDate && payment.payDate.getTime() > deadline.getTime())
      : [];
    const onTimePayments = clientPayments.filter((payment) => !deadline || payment.payDate.getTime() <= deadline.getTime());
    const lateDays = latePayments.reduce((maximum, payment) => {
      const days = Math.ceil((payment.payDate.getTime() - deadline.getTime()) / 86400000);
      return Math.max(maximum, days);
    }, 0);
    const paymentDates = clientPayments.map((payment) => payment.payDate).filter(Boolean).sort((a, b) => a - b);
    let status = 'unpaid';
    if (clientPayments.length) status = latePayments.length ? 'late' : 'paid';
    else if (!overdue) status = 'pending';

    return {
      key,
      contributor: representative.contributor || clientPayments[0]?.contributor || key,
      tax: representative.tax || clientPayments[0]?.tax || '',
      frequency: entries.map((entry) => entry.frequency).filter(Boolean).join(' · '),
      sourceData: representative.sourceData || {},
      sourceRow: representative.sourceRow ?? null,
      sourceSheet: representative.sheetName || '',
      payments: clientPayments,
      amount: total,
      lateAmount: latePayments.reduce((sum, payment) => sum + payment.amount, 0),
      onTimeAmount: onTimePayments.reduce((sum, payment) => sum + payment.amount, 0),
      paymentCount: clientPayments.length,
      firstPayment: paymentDates[0] || null,
      lastPayment: paymentDates[paymentDates.length - 1] || null,
      lateDays,
      status,
    };
  });

  const statusOrder = { late: 0, unpaid: 1, pending: 2, paid: 3 };
  rosterRows.sort((a, b) => {
    const statusDiff = statusOrder[a.status] - statusOrder[b.status];
    if (statusDiff) return statusDiff;
    if (a.status === 'late' && b.status === 'late') return b.lateDays - a.lateDays;
    return b.amount - a.amount || a.contributor.localeCompare(b.contributor, 'es');
  });

  const statusCounts = rosterRows.reduce((counts, row) => {
    counts[row.status] += 1;
    return counts;
  }, { late: 0, unpaid: 0, pending: 0, paid: 0 });
  const totalAmount = periodPayments.reduce((sum, payment) => sum + payment.amount, 0);
  const lateAmount = periodPayments.reduce((sum, payment) => sum + (deadline && payment.payDate.getTime() > deadline.getTime() ? payment.amount : 0), 0);
  const onTimeAmount = totalAmount - lateAmount;
  const paidCount = statusCounts.late + statusCounts.paid;
  const punctualCount = statusCounts.paid;

  const trendKeys = Array.from(new Set(payments.map((payment) => payment.periodKey).filter(Boolean))).sort();
  const trend = trendKeys.map((key) => {
    const rows = getPaymentsForPeriod({ payments, clients, period: key, taxFilter, rosterOnly: rosterLoaded });
    return {
      key,
      label: shortMonthLabel(key),
      amount: rows.reduce((sum, payment) => sum + payment.amount, 0),
      count: rows.length,
    };
  });

  const topClients = Array.from(periodPayments.reduce((map, payment) => {
    if (!map.has(payment.key)) map.set(payment.key, { key: payment.key, name: payment.contributor, amount: 0, count: 0 });
    const item = map.get(payment.key);
    item.amount += payment.amount;
    item.count += 1;
    return map;
  }, new Map()).values())
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 8);

  const byPaymentMonth = Array.from(periodPayments.reduce((map, payment) => {
    const key = monthKey(payment.payDate);
    if (!key) return map;
    map.set(key, (map.get(key) || 0) + payment.amount);
    return map;
  }, new Map()).entries())
    .map(([key, amount]) => ({ key, label: shortMonthLabel(key), amount }))
    .sort((a, b) => a.key.localeCompare(b.key));

  const frequencies = Array.from(new Set(expectedEntries.map((client) => client.frequency).filter(Boolean)));
  const quarterlyWarning = frequencies.some((frequency) => /trimestral|trimestre/i.test(frequency));
  const missingRoster = !rosterLoaded;

  return {
    deadline,
    dueKey,
    overdue,
    rosterLoaded,
    missingRoster,
    expectedCount: rosterRows.length,
    paidCount,
    onTimeClientCount: punctualCount,
    lateCount: statusCounts.late,
    unpaidCount: statusCounts.unpaid,
    pendingCount: statusCounts.pending,
    statusCounts,
    punctuality: paidCount ? (punctualCount / paidCount) * 100 : 0,
    totalAmount,
    lateAmount,
    onTimeAmount,
    periodPayments,
    rosterRows,
    trend,
    topClients,
    byPaymentMonth,
    quarterlyWarning,
    transactions: periodPayments.length,
    matchingClients: paymentMap.size,
    startDate: dateFromKey(`${period}-01`),
  };
}

export function buildPivot({
  payments = [],
  clients = [],
  period = '',
  taxFilter = '',
  dimension = 'period',
  metric = 'amount',
  scope = 'all',
  rosterOnly = false,
  limit = 12,
}) {
  let rows = payments.filter((payment) => matchesTax(payment.taxKey, taxFilter));
  if (scope === 'selected' && period) rows = rows.filter((payment) => payment.periodKey === period);
  if (rosterOnly && clients.length) {
    const rosterByPeriod = new Map();
    rows = rows.filter((payment) => {
      if (!rosterByPeriod.has(payment.periodKey)) {
        rosterByPeriod.set(payment.periodKey, activeRosterMap(clients, payment.periodKey, taxFilter));
      }
      const eligible = rosterByPeriod.get(payment.periodKey);
      return paymentMatchesRoster(payment, eligible.get(payment.key), taxFilter);
    });
  }

  const groups = new Map();
  rows.forEach((payment) => {
    let key = '';
    let label = '';
    if (dimension === 'period') {
      key = payment.periodKey;
      label = monthLabel(payment.periodKey, 'short');
    } else if (dimension === 'paymentMonth') {
      key = monthKey(payment.payDate);
      label = monthLabel(key, 'short');
    } else if (dimension === 'contributor') {
      key = payment.key;
      label = payment.contributor;
    } else if (dimension === 'tax') {
      key = payment.taxKey || payment.tax || 'sin tributo';
      label = payment.tax || 'Sin tributo';
    } else {
      key = payment.dimensions?.[dimension] || 'sin dato';
      label = key === 'sin dato' ? 'Sin dato' : key;
    }
    if (!key) {
      key = 'sin dato';
      label = 'Sin dato';
    }
    if (!groups.has(key)) groups.set(key, { key, label, amount: 0, count: 0, contributors: new Set() });
    const group = groups.get(key);
    group.amount += payment.amount;
    group.count += 1;
    group.contributors.add(payment.key);
  });

  let result = Array.from(groups.values()).map((group) => ({
    ...group,
    value: metric === 'payments' ? group.count : metric === 'contributors' ? group.contributors.size : group.amount,
  }));
  const chronological = dimension === 'period' || dimension === 'paymentMonth';
  result.sort((a, b) => chronological
    ? a.key.localeCompare(b.key)
    : b.value - a.value || a.label.localeCompare(b.label, 'es'));
  const totalGroupCount = result.length;
  if (!chronological && Number(limit) > 0) result = result.slice(0, Number(limit));
  return {
    rows: result,
    totalGroupCount,
    labels: result.map((row) => row.label),
    values: result.map((row) => row.value),
    transactionCount: rows.length,
  };
}

export function uniqueTaxOptions(clients = [], payments = []) {
  const map = new Map();
  [...clients, ...payments].forEach((row) => {
    if (!row.taxKey) return;
    const current = map.get(row.taxKey);
    if (!current || (!current.label.includes(' - ') && row.tax)) map.set(row.taxKey, { key: row.taxKey, label: row.tax || row.taxKey });
  });
  return Array.from(map.values()).sort((a, b) => a.label.localeCompare(b.label, 'es'));
}

export function bestPeriod(payments = []) {
  const counts = new Map();
  payments.forEach((payment) => {
    if (payment.periodKey) counts.set(payment.periodKey, (counts.get(payment.periodKey) || 0) + 1);
  });
  return Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || b[0].localeCompare(a[0]))[0]?.[0] || '';
}
