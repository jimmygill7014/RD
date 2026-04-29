import Field from './Field.jsx';
import DataTable from './DataTable.jsx';
import FamilySection from './FamilySection.jsx';
import EmploymentSection from './EmploymentSection.jsx';
import LiabilitiesSection from './LiabilitiesSection.jsx';
import IncomeSection from './IncomeSection.jsx';
import AssetsSection from './AssetsSection.jsx';
import TaxesExpensesSection from './TaxesExpensesSection.jsx';
import { useStore } from '../store/StoreContext.jsx';

const CUSTOM_RENDERERS = {
  family: FamilySection,
  employment: EmploymentSection,
  liabilities: LiabilitiesSection,
  income: IncomeSection,
  assets: AssetsSection,
  taxesExpenses: TaxesExpensesSection,
};

export default function FormSection({ section }) {
  const { data, update } = useStore();
  const sectionData = data[section.id] || {};
  const flags = data._flags || {};

  const Custom = CUSTOM_RENDERERS[section.id];
  if (Custom) return <Custom section={section} />;

  if (section.customRenderer) {
    return (
      <section className={`form-section theme-${section.colorTheme || 'default'}`}>
        <div className="section-header">
          <h2>{section.title}</h2>
        </div>
        <div className="section-content">
          <p className="section-note">
            (Custom-rendered section — not yet ported. Legacy renderer:{' '}
            <code>{section.customRenderer}</code>)
          </p>
        </div>
      </section>
    );
  }

  const visibleFields = (section.fields || []).filter(f =>
    !f.showIf || flags[f.showIf]
  );

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        {section.note && <p className="section-note">{section.note}</p>}
        {visibleFields.length > 0 && (
          <div className="grid">
            {visibleFields.map(f => (
              <Field
                key={f.key}
                field={f}
                value={sectionData[f.key]}
                onChange={val => update(`${section.id}.${f.key}`, val)}
              />
            ))}
          </div>
        )}
        {section.tables?.map(t => (
          <DataTable key={t.key} sectionId={section.id} tableDef={t} />
        ))}
        {section.conditionalBlocks && (
          <p className="section-note">
            (Conditional blocks not yet ported — coming next.)
          </p>
        )}
      </div>
    </section>
  );
}
