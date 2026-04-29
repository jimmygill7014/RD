import DataTable from './DataTable.jsx';

// Note: when Assets is ported, real-estate rows with remainingLoan > 0 should
// auto-seed read-only rows here (description, payment, amount, interest rate,
// term) tagged with _reKey for live sync to the source. Until then this is a
// plain user-managed list.
const LIABILITIES_TABLE = {
  key: 'items',
  title: '',
  columns: [
    { key: '_reKey', label: '', hidden: true },
    { key: 'description', label: 'Description' },
    { key: 'payment', label: 'Payment (P&I)', type: 'currency', labelInfo: 'Monthly' },
    { key: 'amount', label: 'Amount', type: 'currency' },
    { key: 'interestRate', label: 'Int. Rate', type: 'percent' },
    { key: 'term', label: 'Term (years)' },
  ],
  starterRows: [],
};

export default function LiabilitiesSection({ section }) {
  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        <p className="section-note">
          Known loan liabilities from the Assets section are prefilled below. Add any additional liabilities.
        </p>
        <DataTable sectionId="liabilities" tableDef={LIABILITIES_TABLE} />
      </div>
    </section>
  );
}
