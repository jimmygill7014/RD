import { useStore } from '../store/StoreContext.jsx';
import { getOwnerNameOptions } from '../store/selectors.js';
import { useAutoSourceSync } from '../store/useAutoSourceSync.js';
import DataTable from './DataTable.jsx';

function fullName(first, last) {
  return [(first || '').trim(), (last || '').trim()].filter(Boolean).join(' ');
}

const EMP_READONLY_COLS = ['owner', 'description', 'retirementDate'];
const RE_INCOME_READONLY_COLS = ['owner', 'description', 'annualAmount', 'notes'];

export default function IncomeSection({ section }) {
  const { data } = useStore();
  const ownerOptions = getOwnerNameOptions(data);

  // Employment Income: one auto-row per employed client.
  useAutoSourceSync({
    targetPath: 'income.employment',
    matchKeyField: '_source',
    computeAutoRows: d => {
      const fam = d?.family || {};
      const flags = d?._flags || {};
      const emp = d?.employment || {};
      const rows = [];
      // Status defaults to 'Employed' if not yet persisted (matches the UI).
      const c1Status = emp.client1?.status ?? 'Employed';
      const c2Status = emp.client2?.status ?? 'Employed';
      if (c1Status === 'Employed') {
        rows.push({
          _source: 'client1',
          owner: fullName(fam.client1FirstName, fam.client1LastName),
          description: emp.client1?.employer || '',
          retirementDate: emp.client1?.retirementDate || '',
          _readOnly: EMP_READONLY_COLS,
        });
      }
      if (flags.hasSpouse && c2Status === 'Employed') {
        rows.push({
          _source: 'client2',
          owner: fullName(fam.client2FirstName, fam.client2LastName),
          description: emp.client2?.employer || '',
          retirementDate: emp.client2?.retirementDate || '',
          _readOnly: EMP_READONLY_COLS,
        });
      }
      return rows;
    },
  });

  // Other Income: one auto-row per real-estate asset with rental income.
  useAutoSourceSync({
    targetPath: 'income.other',
    matchKeyField: '_source',
    computeAutoRows: d => {
      const re = Array.isArray(d?.assets?.realEstate) ? d.assets.realEstate : [];
      return re
        .filter(r => r?.incomeEBT && parseFloat(r.incomeEBT) > 0 && r._autoKey)
        .map(r => ({
          _source: `re:${r._autoKey}`,
          owner: r.ownership || '',
          description: r.description || 'Asset Income',
          annualAmount: parseFloat(r.incomeEBT) * 12,
          notes: 'From assets',
          _readOnly: RE_INCOME_READONLY_COLS,
        }));
    },
  });

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
