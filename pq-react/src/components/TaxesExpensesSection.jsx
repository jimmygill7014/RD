import { useStore } from '../store/StoreContext.jsx';
import Field from './Field.jsx';
import DataTable from './DataTable.jsx';

// TODO (with the broader auto-source pass):
//   - Annual Expenses: auto-rows from liabilities.items (payment * 12) and
//     insurance.policies (annualPremium), with description/amount/notes
//     read-only.
//   - Total Expenses computed field — needs the computation engine.

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

function GroupRow({ label }) {
  return (
    <div className="taxexp-group-row">
      <div className="taxexp-group-label">{label}</div>
    </div>
  );
}

export default function TaxesExpensesSection({ section }) {
  const { data, update } = useStore();
  const sectionData = data.taxesExpenses || {};

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

        <GroupRow label="Taxes Paid" />
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
          {/* Total Expenses computed field — deferred until computation engine. */}
        </div>
      </div>
    </section>
  );
}
