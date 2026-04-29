import CurrencyInput from './CurrencyInput.jsx';

export default function TableCell({ column, value, onChange, readOnly = false }) {
  const v = value ?? '';

  if (column.hidden) {
    return <input type="hidden" value={v} readOnly />;
  }

  const handle = readOnly ? () => {} : onChange;
  const inputProps = readOnly ? { readOnly: true } : {};

  switch (column.type) {
    case 'select':
      return (
        <select
          value={v}
          onChange={e => handle(e.target.value)}
          disabled={readOnly}
        >
          <option value="">—</option>
          {(column.options || []).map(o => (
            <option key={o} value={o}>{o}</option>
          ))}
        </select>
      );

    case 'currency':
      return <CurrencyInput value={v} onChange={handle} {...inputProps} />;

    case 'percent':
      return (
        <input
          type="text"
          inputMode="decimal"
          value={v}
          onChange={e => handle(e.target.value)}
          {...inputProps}
        />
      );

    case 'datalist': {
      const listId = `dl-${column.key}-${(column.options || []).join('|')}`;
      return (
        <>
          <input
            type="text"
            list={listId}
            value={v}
            onChange={e => handle(e.target.value)}
            {...inputProps}
          />
          <datalist id={listId}>
            {(column.options || []).map(o => <option key={o} value={o} />)}
          </datalist>
        </>
      );
    }

    case 'date':
    case 'number':
      return (
        <input
          type={column.type}
          value={v}
          onChange={e => handle(e.target.value)}
          {...inputProps}
        />
      );

    default:
      return (
        <input
          type="text"
          value={v}
          onChange={e => handle(e.target.value)}
          {...inputProps}
        />
      );
  }
}
