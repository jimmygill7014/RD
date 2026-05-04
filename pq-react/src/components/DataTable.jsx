import { useStore } from '../store/StoreContext.jsx';
import { genUid } from '../store/uid.js';
import TableCell from './TableCell.jsx';
import { useEffect } from 'react';

function formatCommas(n) {
  if (n == null || isNaN(n)) return '';
  return n % 1 === 0
    ? n.toLocaleString('en-US')
    : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function sumColumn(rows, colKey) {
  return rows.reduce((sum, r) => {
    const raw = r?.[colKey];
    if (raw == null || raw === '') return sum;
    const n = parseFloat(String(raw).replace(/,/g, ''));
    return isNaN(n) ? sum : sum + n;
  }, 0);
}

export default function DataTable({ sectionId, tableDef, noSeed = false }) {
  const { data, update } = useStore();
  const path = `${sectionId}.${tableDef.key}`;
  const stored = data?.[sectionId]?.[tableDef.key];
  const rows = Array.isArray(stored) ? stored.filter(r => r != null) : null;

  // Seed rows on first render if storage is empty.
  // Skipped when the parent manages rows externally (noSeed = true).
  useEffect(() => {
    if (noSeed) return;
    if (rows && rows.length > 0) return;
    const seed = (tableDef.starterRows && tableDef.starterRows.length)
      ? tableDef.starterRows.map(r => ({ ...r, _uid: genUid() }))
      : [{ _uid: genUid() }];
    update(path, seed);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const current = rows && rows.length > 0 ? rows : [];

  const setCell = (rowIdx, colKey, value) => {
    const next = current.map((r, i) => i === rowIdx ? { ...r, [colKey]: value } : r);
    update(path, next);
  };

  const addRow = () => update(path, [...current, { _uid: genUid() }]);

  const removeRow = rowIdx => {
    const next = current.filter((_, i) => i !== rowIdx);
    update(path, noSeed ? next : (next.length ? next : [{ _uid: genUid() }]));
  };

  const showTotals = !!tableDef.showTotals;

  return (
    <div>
      <div className="table-tools">
        {tableDef.title && (
          <div className={'table-title' + (tableDef.titleClass ? ' ' + tableDef.titleClass : '')}>
            {tableDef.titleClass && <span className="table-title-dot" aria-hidden="true" />}
            {tableDef.title}
          </div>
        )}
        <button type="button" className="btn-link" onClick={addRow}>
          + Add Row
        </button>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              {tableDef.columns.map(c => (
                <th
                  key={c.key}
                  className={[
                    c.type === 'percent' ? 'th-percent' : '',
                    c.colClass || '',
                  ].filter(Boolean).join(' ') || undefined}
                  style={c.hidden ? { display: 'none' } : undefined}
                >
                  {c.label}
                  {c.labelInfo && (
                    <span
                      className="th-info"
                      tabIndex={0}
                      title={c.labelInfo}
                      aria-label={c.labelInfo}
                    >
                      i
                    </span>
                  )}
                </th>
              ))}
              <th style={{ width: 40 }} aria-label="Actions" />
            </tr>
          </thead>
          <tbody>
            {current.map((row, rowIdx) => (
              <tr key={row._uid || row._autoKey || row._source || row._reKey || rowIdx}>
                {tableDef.columns.map(col => (
                  <td
                    key={col.key}
                    style={col.hidden ? { display: 'none' } : undefined}
                  >
                    <TableCell
                      column={col}
                      value={row[col.key]}
                      readOnly={Array.isArray(row._readOnly) && row._readOnly.includes(col.key)}
                      onChange={val => setCell(rowIdx, col.key, val)}
                    />
                  </td>
                ))}
                <td>
                  <button
                    type="button"
                    className="btn-icon"
                    aria-label="Remove row"
                    onClick={() => removeRow(rowIdx)}
                  >
                    ×
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
          {showTotals && current.length > 0 && (
            <tfoot>
              <tr>
                {tableDef.columns.map((col, ci) => {
                  if (col.hidden) return <td key={col.key} style={{ display: 'none' }} />;
                  if (ci === 0 && col.type !== 'currency') {
                    return <td key={col.key} style={{ fontWeight: 600 }}>Total</td>;
                  }
                  if (col.type === 'currency') {
                    return (
                      <td key={col.key} className="tfoot-total">
                        {formatCommas(sumColumn(current, col.key))}
                      </td>
                    );
                  }
                  return <td key={col.key} />;
                })}
                <td />
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}
