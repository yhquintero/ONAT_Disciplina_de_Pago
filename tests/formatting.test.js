import test from 'node:test';
import assert from 'node:assert/strict';
import { formatCompactMoney, formatMoney } from '../src/formatting.js';

test('muestra el símbolo monetario antes de la cifra con separadores solicitados', () => {
  assert.equal(formatMoney(1_000_000), '$ 1,000,000.00');
  assert.equal(formatMoney(1234.5), '$ 1,234.50');
  assert.equal(formatMoney('no numérico'), '$ 0.00');
});

test('usa el mismo prefijo en etiquetas compactas de gráficos', () => {
  assert.match(formatCompactMoney(1_000_000), /^\$ /);
  assert.doesNotMatch(formatCompactMoney(1_000_000), / \$$/);
});
