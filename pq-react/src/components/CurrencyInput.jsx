import { useState } from 'react';

function formatCommas(val) {
  const n = parseFloat(String(val).replace(/,/g, ''));
  if (isNaN(n)) return '';
  return n % 1 === 0
    ? n.toLocaleString('en-US')
    : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseCommas(val) {
  return parseFloat(String(val).replace(/,/g, ''));
}

export default function CurrencyInput({ value, onChange, ...rest }) {
  const [focused, setFocused] = useState(false);

  const display = focused
    ? value === '' || value == null ? '' : String(value).replace(/,/g, '')
    : value === '' || value == null ? '' : formatCommas(value);

  return (
    <input
      type="text"
      inputMode="decimal"
      value={display}
      onFocus={() => setFocused(true)}
      onBlur={() => setFocused(false)}
      onChange={e => {
        const raw = e.target.value;
        const n = parseCommas(raw);
        onChange(raw === '' ? '' : isNaN(n) ? raw : n);
      }}
      {...rest}
    />
  );
}
