import { useStore } from '../store/StoreContext.jsx';
import { getOwnerNameOptions } from '../store/selectors.js';
import DataTable from './DataTable.jsx';

// TODO: cross-section auto-source (deferred to a follow-up):
//   - Employment Income: prepend rows from employment.client1/client2 when
//     status === 'Employed', sourced via _source: 'client1'|'client2', with
//     owner/description/retirementDate read-only.
//   - Other Income: prepend rows from assets.realEstate where incomeEBT > 0,
//     with owner/description/annualAmount/notes read-only.
// Until then, all four tables are plain user-managed lists.

export default function IncomeSection({ section }) {
  const { data } = useStore();
  const ownerOptions = getOwnerNameOptions(data);

  const tables = [
    {
      key: 'employment',
      title: 'Employment Income',
      columns: [
        { key: '_source', label: '', hidden: true },
        { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
        { key: 'description', label: 'Description' },
        { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
        { key: 'retirementDate', label: 'Retirement Date', type: 'date' },
        { key: 'cola', label: 'COLA %', type: 'percent' },
      ],
    },
    {
      key: 'socialSecurity',
      title: 'Social Security',
      columns: [
        { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
        { key: 'startingAge', label: 'Starting Age', type: 'number' },
        { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
        { key: 'fullRetAge', label: 'Full Ret. Age', type: 'number' },
        { key: 'cola', label: 'COLA %', type: 'percent' },
      ],
    },
    {
      key: 'pension',
      title: 'Pension',
      columns: [
        { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
        { key: 'startDate', label: 'Start Date', type: 'date' },
        { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
        { key: 'survivorBenefit', label: 'Survivor Benefit', type: 'percent' },
        { key: 'cola', label: 'COLA %', type: 'percent' },
      ],
    },
    {
      key: 'other',
      title: 'Other Income',
      columns: [
        { key: '_source', label: '', hidden: true },
        { key: 'owner', label: 'Owner', type: 'datalist', options: ownerOptions },
        { key: 'description', label: 'Description' },
        { key: 'annualAmount', label: 'Annual Amount', type: 'currency' },
        { key: 'startDate', label: 'Start Date', type: 'date' },
        { key: 'endDate', label: 'End Date', type: 'date' },
        { key: 'cola', label: 'COLA %', type: 'percent' },
        { key: 'notes', label: 'Notes' },
      ],
    },
  ];

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        {tables.map(t => (
          <DataTable key={t.key} sectionId="income" tableDef={t} />
        ))}
        {/* TODO: Total Income computed field — needs the computation engine,
            which we haven't ported yet. */}
      </div>
    </section>
  );
}
