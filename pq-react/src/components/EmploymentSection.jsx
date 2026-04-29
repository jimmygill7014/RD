import { useStore } from '../store/StoreContext.jsx';
import Field from './Field.jsx';

const STATUS_OPTIONS = ['Employed', 'Retired', 'Not Employed'];

const EMPLOYED_FIELDS = [
  { key: 'jobTitle', label: 'Job Title', width: 'medium' },
  { key: 'employer', label: 'Employer', width: 'medium' },
  { key: 'years', label: '# of Years', type: 'number', width: 'field' },
  { key: 'retirementDate', label: 'Retirement Date', type: 'date', width: 'medium' },
  { key: 'retirementAge', label: 'Retirement Age', type: 'number', width: 'field' },
  { key: 'businessType', label: 'Type of Business', width: 'medium' },
];

const NOT_EMPLOYED_FIELDS = [
  { key: 'lastJob', label: 'Last Job Title', width: 'medium' },
  { key: 'lastEmployer', label: 'Last Employer', width: 'medium' },
  { key: 'yearsWorked', label: 'Years Worked', type: 'number', width: 'field' },
];

function EmploymentSubsection({ title, clientKey, marginTop }) {
  const { data, update } = useStore();
  const employment = data.employment || {};
  const personData = employment[clientKey] || {};
  const status = personData.status ?? 'Employed';

  const fields = status === 'Employed' ? EMPLOYED_FIELDS : NOT_EMPLOYED_FIELDS;

  return (
    <>
      <h3 className="subsection-title" style={marginTop ? { marginTop } : undefined}>
        {title}
      </h3>
      <div className="radio-group">
        {STATUS_OPTIONS.map(opt => (
          <label key={opt} className="radio-item">
            <input
              type="radio"
              name={`employment.${clientKey}.status`}
              value={opt}
              checked={status === opt}
              onChange={() => update(`employment.${clientKey}.status`, opt)}
            />
            {opt}
          </label>
        ))}
      </div>
      <div className="grid">
        {fields.map(f => (
          <Field
            key={f.key}
            field={f}
            value={personData[f.key]}
            onChange={val => update(`employment.${clientKey}.${f.key}`, val)}
          />
        ))}
      </div>
    </>
  );
}

export default function EmploymentSection({ section }) {
  const { data } = useStore();
  const hasSpouse = !!data?._flags?.hasSpouse;

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        <EmploymentSubsection title="Client 1 Employment" clientKey="client1" />
        {hasSpouse && (
          <EmploymentSubsection
            title="Spouse Employment"
            clientKey="client2"
            marginTop={16}
          />
        )}
      </div>
    </section>
  );
}
