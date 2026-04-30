// Pure selector helpers that read from the store data shape.

function toNumber(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function fullName(first, last) {
  return [(first || '').trim(), (last || '').trim()].filter(Boolean).join(' ');
}

function isFutureDate(dateStr) {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return false;
  return d.getFullYear() > new Date().getFullYear();
}

export function getTotalIncome(data) {
  const income = data?.income || {};
  const family = data?.family || {};
  const c1Age = toNumber(family.client1Age);
  const c2Age = toNumber(family.client2Age);
  const c1Name = fullName(family.client1FirstName, family.client1LastName);
  const c2Name = fullName(family.client2FirstName, family.client2LastName);

  let total = 0;

  // Employment — always included
  (income.employment || []).forEach(r => {
    if (r) total += toNumber(r.annualAmount);
  });

  // Social Security — exclude if startingAge > current age of owner
  (income.socialSecurity || []).forEach(r => {
    if (!r) return;
    const amt = toNumber(r.annualAmount);
    const startAge = toNumber(r.startingAge);
    let ownerAge = c1Age;
    if (c2Name && r.owner === c2Name) ownerAge = c2Age;
    if (startAge > 0 && startAge > ownerAge) return;
    total += amt;
  });

  // Pension — exclude future-dated
  (income.pension || []).forEach(r => {
    if (!r) return;
    if (isFutureDate(r.startDate)) return;
    total += toNumber(r.annualAmount);
  });

  // Other — exclude future-dated
  (income.other || []).forEach(r => {
    if (!r) return;
    if (isFutureDate(r.startDate)) return;
    total += toNumber(r.annualAmount);
  });

  return total;
}

export function getTotalExpenses(data) {
  const te = data?.taxesExpenses || {};
  let total = 0;
  (te.expenses || []).forEach(r => {
    if (r) total += toNumber(r.amount);
  });
  total += toNumber(te.livingExpenses);
  return total;
}

export function getTotalTaxesPaid(data) {
  const te = data?.taxesExpenses || {};
  return toNumber(te.federalTax) + toNumber(te.stateTax) + toNumber(te.ficaTax);
}

export function getEffectiveTaxRate(data) {
  const te = data?.taxesExpenses || {};
  const taxableInc = toNumber(te.taxableIncome);
  const totalTax = getTotalTaxesPaid(data);
  if (taxableInc <= 0 || totalTax <= 0) return null;
  return (totalTax / taxableInc) * 100;
}

function sumKey(rows, key) {
  if (!Array.isArray(rows)) return 0;
  return rows.reduce((s, r) => s + (r ? toNumber(r[key]) : 0), 0);
}

const ASSET_TABLES = ['realEstate', 'taxDeferred', 'roth', 'taxable', 'cashCd', 'plan529', 'hsa', 'businessOther'];

export function getDashboardMetrics(data) {
  const assets = data?.assets || {};
  const liabilities = data?.liabilities || {};
  const taxesExpenses = data?.taxesExpenses || {};

  const totalAssets = ASSET_TABLES.reduce((s, k) => s + sumKey(assets[k], 'marketValue'), 0);
  const realEstateAssets = sumKey(assets.realEstate, 'marketValue');
  const businessOtherAssets = sumKey(assets.businessOther, 'marketValue');
  const investableAssets = totalAssets - realEstateAssets - businessOtherAssets;

  // Personal additions only (savings); excludes company match.
  const totalSavings = ASSET_TABLES.reduce((s, k) => s + sumKey(assets[k], 'personalAdditions'), 0);

  const totalLiabilities = sumKey(liabilities.items, 'amount');
  const totalNetWorth = totalAssets - totalLiabilities;

  const totalIncome = getTotalIncome(data);
  const totalTax = getTotalTaxesPaid(data);

  // Split annual expenses: liability-sourced vs everything else (+ separate
  // livingExpenses field). Legacy used a hidden _source = 'liability' marker;
  // our auto-source keys are 'liability:<id>', so prefix-match.
  let liabilityPayments = 0;
  let nonLiabilityExpenses = 0;
  const expenses = Array.isArray(taxesExpenses.expenses) ? taxesExpenses.expenses : [];
  expenses.forEach(r => {
    if (!r) return;
    const amt = toNumber(r.amount);
    if (typeof r._source === 'string' && r._source.startsWith('liability')) {
      liabilityPayments += amt;
    } else {
      nonLiabilityExpenses += amt;
    }
  });
  nonLiabilityExpenses += toNumber(taxesExpenses.livingExpenses);

  const cashFlow = totalIncome - totalSavings - liabilityPayments - totalTax - nonLiabilityExpenses;

  const dti = totalIncome > 0 ? (totalLiabilities / totalIncome) * 100 : 0;
  const savingsRatio = totalIncome > 0 ? (totalSavings / totalIncome) * 100 : 0;

  // Tax triangle buckets
  const taxFree = sumKey(assets.roth, 'marketValue') + sumKey(assets.hsa, 'marketValue');
  const taxDeferred = sumKey(assets.taxDeferred, 'marketValue');
  const taxable = sumKey(assets.taxable, 'marketValue') + sumKey(assets.cashCd, 'marketValue');

  return {
    totalAssets, realEstateAssets, businessOtherAssets, investableAssets,
    totalLiabilities, totalNetWorth,
    totalIncome, totalSavings, totalTax,
    liabilityPayments, nonLiabilityExpenses, cashFlow,
    dti, savingsRatio,
    taxFree, taxDeferred, taxable,
  };
}

export function formatDollars(n) {
  if (n == null || isNaN(n)) return '$0';
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(n);
  return sign + '$' + (abs % 1 === 0
    ? abs.toLocaleString('en-US')
    : abs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
}

export function getOwnerNameOptions(data) {
  const family = data?.family || {};
  const flags = data?._flags || {};

  const c1 = [family.client1FirstName, family.client1LastName]
    .map(s => (s || '').trim())
    .filter(Boolean)
    .join(' ');
  const c2 = [family.client2FirstName, family.client2LastName]
    .map(s => (s || '').trim())
    .filter(Boolean)
    .join(' ');

  const opts = [];
  if (c1) opts.push(c1);
  if ((flags.hasSpouse || family.client2FirstName || family.client2LastName) && c2) {
    opts.push(c2);
  }
  if (c1 && c2) opts.push('Joint');
  opts.push('Trust');
  return opts;
}
