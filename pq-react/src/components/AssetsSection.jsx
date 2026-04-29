import { useEffect } from 'react';
import { useStore } from '../store/StoreContext.jsx';
import { getOwnerNameOptions } from '../store/selectors.js';
import DataTable from './DataTable.jsx';

const RE_CHECK_DEFS = [
  { key: 'ownHome',       label: 'Own your home?',         desc: 'Primary Residence' },
  { key: 'secondaryHome', label: 'Secondary home(s)?',     desc: 'Secondary Residence', legacyDescs: ['Secondary Home'] },
  { key: 'additionalRE',  label: 'Additional real estate?', desc: 'Additional Real Estate' },
];

const INV_CHECK_DEFS = [
  { key: 'taxDeferred', label: 'Tax-Deferred Retirement' },
  { key: 'roth',        label: 'Tax-Free Roth' },
  { key: 'taxable',     label: 'Taxable Non-Retirement' },
  { key: 'cashCd',      label: "Cash & CD's" },
  { key: 'plan529',     label: '529 Plans' },
  { key: 'hsa',         label: 'HSA' },
];

const RE_DATA_KEYS = [
  'marketValue', 'remainingLoan', 'interestRate', 'term', 'payment',
  'yearAcquired', 'costBasis', 'ownership', 'incomeEBT',
];

function rowHasData(row) {
  return RE_DATA_KEYS.some(k => {
    const v = row?.[k];
    if (v == null) return false;
    if (typeof v === 'string') return v.trim() !== '';
    return true;
  });
}

function CheckboxGroup({ defs, checks, onToggle }) {
  return (
    <div className="checkbox-group">
      {defs.map(d => (
        <label key={d.key} className="checkbox-item">
          <input
            type="checkbox"
            checked={!!checks[d.key]}
            onChange={e => onToggle(d.key, e.target.checked)}
          />
          {d.label}
        </label>
      ))}
    </div>
  );
}

export default function AssetsSection({ section }) {
  const { data, update } = useStore();
  const ownershipOptions = getOwnerNameOptions(data);
  const checks = data._assetChecks || {};

  const setCheck = (key, value) => update(`_assetChecks.${key}`, value);

  // Reconcile real-estate auto-rows with checkbox state.
  // For each checked box: ensure a row exists with that _autoKey (description tied to box).
  // For each unchecked box: remove the row IF it has no user data beyond description.
  useEffect(() => {
    const existing = Array.isArray(data?.assets?.realEstate) ? data.assets.realEstate : [];
    let merged = existing.filter(r => r != null);
    let changed = false;

    RE_CHECK_DEFS.forEach(def => {
      const idx = merged.findIndex(r => r?._autoKey === def.key);
      if (checks[def.key]) {
        if (idx === -1) {
          merged.push({ _autoKey: def.key, description: def.desc });
          changed = true;
        } else if (merged[idx].description !== def.desc) {
          merged[idx] = { ...merged[idx], description: def.desc };
          changed = true;
        }
      } else if (idx !== -1 && !rowHasData(merged[idx])) {
        merged.splice(idx, 1);
        changed = true;
      }
    });

    if (changed) update('assets.realEstate', merged);
  }, [checks.ownHome, checks.secondaryHome, checks.additionalRE, data?.assets?.realEstate, update]);

  const reColumns = [
    { key: '_autoKey',      label: '',               hidden: true },
    { key: 'description',   label: 'Description',    colClass: 'col-wide' },
    { key: 'marketValue',   label: 'Market Value',   type: 'currency' },
    { key: 'remainingLoan', label: 'Remaining Loan', type: 'currency' },
    { key: 'interestRate',  label: 'Int. Rate',      type: 'percent' },
    { key: 'term',          label: 'Term',           colClass: 'col-narrow', labelInfo: 'In years' },
    { key: 'payment',       label: 'Payment (P&I)',  type: 'currency', labelInfo: 'Monthly' },
    { key: 'yearAcquired',  label: 'Year Acquired',  colClass: 'col-narrow' },
    { key: 'costBasis',     label: 'Cost Basis',     type: 'currency' },
    { key: 'ownership',     label: 'Ownership',      type: 'datalist', options: ownershipOptions },
    { key: 'incomeEBT',     label: 'Income (EBT)',   type: 'currency' },
  ];

  const realEstateRows = Array.isArray(data?.assets?.realEstate) ? data.assets.realEstate : [];
  const anyREChecked = RE_CHECK_DEFS.some(d => checks[d.key]);
  const showRETable = anyREChecked || realEstateRows.length > 0;

  const investmentTables = {
    taxDeferred: {
      key: 'taxDeferred', title: 'Tax-Deferred Retirement', titleClass: 'tone-tax-deferred', showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ],
    },
    roth: {
      key: 'roth', title: 'Tax-Free Roth', titleClass: 'tone-tax-free', showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'companyMatch', label: 'Company Match', type: 'currency' },
        { key: 'type', label: 'Type' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ],
    },
    taxable: {
      key: 'taxable', title: 'Taxable Non-Retirement', titleClass: 'tone-taxable', showTotals: true,
      columns: [
        { key: 'custodian', label: 'Custodian' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'personalAdditions', label: 'Personal Additions', type: 'currency', labelInfo: 'Yearly' },
        { key: 'costBasis', label: 'Cost Basis', type: 'currency' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
        { key: 'variableFixed', label: 'Variable / Fixed' },
        { key: 'issueDate', label: 'Issue Date', type: 'date' },
        { key: 'beneficiary', label: 'Beneficiary', type: 'datalist', options: ownershipOptions },
      ],
    },
    cashCd: {
      key: 'cashCd', title: "Cash & CD's", showTotals: true,
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ],
    },
    plan529: {
      key: 'plan529', title: '529 Plans', showTotals: true,
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ],
    },
    hsa: {
      key: 'hsa', title: 'HSA Accounts', showTotals: true,
      columns: [
        { key: 'description', label: 'Description' },
        { key: 'marketValue', label: 'Market Value', type: 'currency' },
        { key: 'interestRate', label: 'Interest Rate' },
        { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
      ],
    },
  };

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        {/* --- Real Estate --- */}
        <h3 className="subsection-title">Real Estate</h3>
        <CheckboxGroup defs={RE_CHECK_DEFS} checks={checks} onToggle={setCheck} />
        {showRETable && (
          <DataTable
            sectionId="assets"
            tableDef={{
              key: 'realEstate',
              title: '',
              showTotals: true,
              columns: reColumns,
            }}
            noSeed
          />
        )}

        {/* --- Business and Other --- */}
        <h3 className="subsection-title" style={{ marginTop: 18 }}>Business and Other</h3>
        <DataTable
          sectionId="assets"
          tableDef={{
            key: 'businessOther',
            title: '',
            showTotals: true,
            columns: [
              { key: 'description', label: 'Description' },
              { key: 'marketValue', label: 'Market Value', type: 'currency' },
              { key: 'costBasis', label: 'Cost Basis', type: 'currency' },
              { key: 'ownership', label: 'Ownership', type: 'datalist', options: ownershipOptions },
            ],
          }}
        />

        {/* --- Investment Accounts --- */}
        <h3 className="subsection-title" style={{ marginTop: 18 }}>Investment Accounts</h3>
        <CheckboxGroup defs={INV_CHECK_DEFS} checks={checks} onToggle={setCheck} />
        {INV_CHECK_DEFS.map(def =>
          checks[def.key] ? (
            <DataTable
              key={def.key}
              sectionId="assets"
              tableDef={investmentTables[def.key]}
            />
          ) : null
        )}
      </div>
    </section>
  );
}
