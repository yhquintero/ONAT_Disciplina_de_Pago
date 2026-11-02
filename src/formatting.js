const currencyFormat = new Intl.NumberFormat('en-US', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const compactCurrencyFormat = new Intl.NumberFormat('en-US', {
  notation: 'compact',
  maximumFractionDigits: 1,
});

function finiteNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

export function formatMoney(value) {
  return `$ ${currencyFormat.format(finiteNumber(value))}`;
}

export function formatCompactMoney(value) {
  return `$ ${compactCurrencyFormat.format(finiteNumber(value))}`;
}
