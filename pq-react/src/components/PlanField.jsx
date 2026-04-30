export default function PlanField({ field, value, onChange }) {
  if (field.type === 'checkbox') {
    return (
      <div className="plan-field-row plan-field-row--check">
        <input
          type="checkbox"
          id={`plan-${field.key}`}
          checked={value === true || value === 'true'}
          onChange={e => onChange(e.target.checked)}
        />
        <label htmlFor={`plan-${field.key}`}>{field.label}</label>
      </div>
    );
  }

  return (
    <div className="plan-field-row">
      <label htmlFor={`plan-${field.key}`} className="plan-field-label">{field.label}</label>
      {renderControl(field, value, onChange)}
    </div>
  );
}

function renderControl(field, value, onChange) {
  const v = value ?? '';
  const id = `plan-${field.key}`;
  const cls = 'plan-field-input';

  if (field.type === 'select') {
    return (
      <select
        id={id}
        className={cls}
        value={v}
        onChange={e => onChange(e.target.value)}
      >
        {(field.options || []).map(opt => (
          <option key={opt} value={opt}>{opt || '— Select —'}</option>
        ))}
      </select>
    );
  }

  if (field.type === 'textarea') {
    return (
      <textarea
        id={id}
        className="plan-field-input plan-field-ta"
        rows={3}
        value={v}
        onChange={e => onChange(e.target.value)}
      />
    );
  }

  return (
    <input
      type={field.type === 'date' ? 'date' : 'text'}
      id={id}
      className={cls}
      value={v}
      onChange={e => onChange(e.target.value)}
    />
  );
}
