import DataTable from './DataTable.jsx';
import { useAutoSourceSync } from '../store/useAutoSourceSync.js';

const RE_READONLY_COLS = ['description', 'payment', 'amount', 'interestRate', 'term'];

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
  useAutoSourceSync({
    targetPath: 'liabilities.items',
    matchKeyField: '_reKey',
    computeAutoRows: data => {
      const re = Array.isArray(data?.assets?.realEstate) ? data.assets.realEstate : [];
      return re
        .filter(r => r?.remainingLoan && parseFloat(r.remainingLoan) > 0 && r._autoKey)
        .map(r => ({
          _reKey: r._autoKey,
          description: (r.description || 'RE Loan') + ' Mortgage',
          payment: r.payment || '',
          amount: r.remainingLoan,
          interestRate: r.interestRate || '',
          term: r.term || '',
          _readOnly: RE_READONLY_COLS,
        }));
    },
  });

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
