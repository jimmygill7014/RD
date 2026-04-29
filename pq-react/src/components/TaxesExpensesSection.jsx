import { useStore } from '../store/StoreContext.jsx';
import { useAutoSourceSync } from '../store/useAutoSourceSync.js';
import {
  getTotalExpenses,
  getTotalTaxesPaid,
  getEffectiveTaxRate,
  formatDollars,
} from '../store/selectors.js';
import Field from './Field.jsx';
import DataTable from './DataTable.jsx';

const EXPENSE_READONLY_COLS = ['description', 'amount', 'notes'];

const PRIOR_INCOME_FIELDS = [
  { key: 'taxableIncome', label: 'Taxable Income', type: 'currency', width: 'medium' },
  { key: 'standardItemDeduction', label: 'Std/Item Deduction', type: 'currency', width: 'medium' },
  { key: 'capLossCarryForward', label: 'Cap Loss Carry Forward', type: 'currency', width: 'medium' },
];

const PRIOR_TAXES_FIELDS = [
  { key: 'federalTax', label: 'Federal', type: 'currency', width: 'medium' },
  { key: 'stateTax', label: 'State', type: 'currency', width: 'medium' },
  { key: 'ficaTax', label: 'FICA', type: 'currency', width: 'medium' },
];

const EXPENSES_TABLE = {
  key: 'expenses',
  title: '',
  columns: [
    { key: '_source', label: '', hidden: true },
    { key: 'description', label: 'Description' },
    { key: 'amount', label: 'Annual Amount', type: 'currency' },
    { key: 'startDate', label: 'Start Date', type: 'date' },
    { key: 'endDate', label: 'End Date', type: 'date' },
    { key: 'cola', label: 'COLA %', type: 'percent' },
    { key: 'notes', label: 'Notes' },
  ],
};

function GroupRow({ label, totalText, hintText }) {
  return (
    <div className="taxexp-group-row">
      <div className="taxexp-group-label">{label}</div>
      {totalText != null && (
        <div className="taxexp-total-badge">
          <input type="text" className="taxexp-total-val" value={totalText} readOnly />
          {hintText && <span className="taxexp-total-hint">{hintText}</span>}
        </div>
      )}
    </div>
  );
}

export default function TaxesExpensesSection({ section }) {
  const { data, update } = useStore();
  const sectionData = data.taxesExpenses || {};

  // Annual Expenses auto-rows: liabilities (monthly payment * 12) + insurance
  // (annual premium). Both share the same _source-keyed shape.
  useAutoSourceSync({
    targetPath: 'taxesExpenses.expenses',
    matchKeyField: '_source',
    computeAutoRows: d => {
      const rows = [];
      const liabs = Array.isArray(d?.liabilities?.items) ? d.liabilities.items : [];
      liabs.forEach((l, i) => {
        if (l?.payment && parseFloat(l.payment) > 0) {
          rows.push({
            _source: `liability:${l._reKey || i}`,
            description: l.description || 'Loan Payment',
            amount: parseFloat(l.payment) * 12,
            notes: 'From liabilities',
            _readOnly: EXPENSE_READONLY_COLS,
          });
        }
      });
      const policies = Array.isArray(d?.insurance?.policies) ? d.insurance.policies : [];
      policies.forEach((p, i) => {
        if (p?.annualPremium && parseFloat(p.annualPremium) > 0) {
          rows.push({
            _source: `insurance:${i}`,
            description: 'Insurance Premium - ' + (p.company || 'Unknown'),
            amount: parseFloat(p.annualPremium),
            notes: 'From insurance',
            _readOnly: EXPENSE_READONLY_COLS,
          });
        }
      });
      return rows;
    },
  });

  const renderField = f => (
    <Field
      key={f.key}
      field={f}
      value={sectionData[f.key]}
      onChange={val => update(`taxesExpenses.${f.key}`, val)}
    />
  );

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        {/* --- Prior Year Tax Information --- */}
        <h3 className="subsection-title">Prior Year Tax Information</h3>

        <GroupRow label="Income & Deductions" />
        <div className="grid">{PRIOR_INCOME_FIELDS.map(renderField)}</div>

        <GroupRow
          label="Taxes Paid"
          totalText={formatDollars(getTotalTaxesPaid(data))}
          hintText={(() => {
            const rate = getEffectiveTaxRate(data);
            return rate != null
              ? `${rate.toFixed(1)}% effective rate on taxable income`
              : 'Federal + State + FICA';
          })()}
        />
        <div className="grid">{PRIOR_TAXES_FIELDS.map(renderField)}</div>

        {/* --- Annual Expenses --- */}
        <h3 className="subsection-title" style={{ marginTop: 14 }}>Annual Expenses</h3>
        <DataTable sectionId="taxesExpenses" tableDef={EXPENSES_TABLE} />

        <div className="grid" style={{ marginTop: 10 }}>
          <Field
            field={{ key: 'livingExpenses', label: 'Living Expenses', type: 'currency', width: 'medium' }}
            value={sectionData.livingExpenses}
            onChange={val => update('taxesExpenses.livingExpenses', val)}
          />
          <label className="field-medium field-computed field-total">
            <div className="label-row">
              <span className="label">Total Expenses</span>
            </div>
            <input type="text" value={formatDollars(getTotalExpenses(data))} readOnly />
          </label>
        </div>
      </div>
    </section>
  );
}
