import CurrencyInput from './CurrencyInput.jsx';

const WIDTH_CLASS = {
  wide: 'field-wide',
  medium: 'field-medium',
  field: 'field',
  'field-2': 'field-2',
};

export default function Field({ field, value, onChange }) {
  let cls = WIDTH_CLASS[field.width] || 'field';
  if (field.type === 'textarea') cls = 'field-full';
  if (field.type === 'computed') cls += ' field-computed';
  if (field.emphasis === 'total') cls += ' field-total';

  // Use <div> for fields that contain multiple form controls (multiselect,
  // future radio groups). A <label> wrapper would forward stray clicks to
  // the first control inside.
  const Wrapper = field.type === 'multiselect' ? 'div' : 'label';

  return (
    <Wrapper className={cls}>
      <div className="label-row">
        <span className="label">{field.label}</span>
      </div>
      {renderControl(field, value, onChange)}
    </Wrapper>
  );
}

function renderControl(field, value, onChange) {
  const v = value ?? '';

  switch (field.type) {
    case 'textarea':
      return (
        <textarea
          value={v}
          onChange={e => onChange(e.target.value)}
          required={field.required || undefined}
        />
      );

    case 'select':
      return (
        <select
          value={v}
          onChange={e => onChange(e.target.value)}
          required={field.required || undefined}
        >
          <option value=""></option>
          {(field.options || []).map(opt => (
            <option key={opt} value={opt}>{opt}</option>
          ))}
        </select>
      );

    case 'multiselect': {
      const selected = Array.isArray(value) ? value : [];
      const toggle = opt => {
        const next = selected.includes(opt)
          ? selected.filter(x => x !== opt)
          : [...selected, opt];
        onChange(next);
      };
      return (
        <div className="multiselect-group">
          {(field.options || []).map(opt => (
            <label key={opt} className="multiselect-item">
              <input
                type="checkbox"
                checked={selected.includes(opt)}
                onChange={() => toggle(opt)}
              />
              <span>{opt}</span>
            </label>
          ))}
        </div>
      );
    }

    case 'currency':
      return <CurrencyInput value={v} onChange={onChange} />;

    case 'computed':
      return <input type="text" value={v} readOnly />;

    case 'date':
    case 'number':
    case 'email':
      return (
        <input
          type={field.type}
          value={v}
          onChange={e => onChange(e.target.value)}
          required={field.required || undefined}
        />
      );

    default:
      return (
        <input
          type="text"
          value={v}
          onChange={e => onChange(e.target.value)}
          required={field.required || undefined}
        />
      );
  }
}
