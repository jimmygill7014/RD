import { useEffect } from 'react';
import { useStore } from './StoreContext.jsx';

/**
 * Generic cross-section auto-source sync.
 *
 * Watches the live store, computes a list of "auto-rows" via `computeAutoRows`,
 * and merges them into the target table at `targetPath`:
 *   - Auto-rows are matched against existing rows via `matchKeyField`
 *     (typically '_source' or '_reKey').
 *   - When an auto-row matches an existing row, user-entered fields are
 *     preserved; computed (read-only) fields are refreshed from source.
 *   - Auto-rows whose source no longer exists are removed.
 *   - Manual rows (no `matchKeyField`) pass through unchanged.
 *   - Writes only when the merged array differs from current state, so this
 *     does not loop.
 *
 * Auto-rows MUST include a `_readOnly` array listing the read-only column keys.
 *
 * @param {object}   args
 * @param {string}   args.targetPath       — e.g. 'liabilities.items'
 * @param {Function} args.computeAutoRows  — (data) => Array<row>
 * @param {string}   args.matchKeyField    — '_source' | '_reKey'
 */
export function useAutoSourceSync({ targetPath, computeAutoRows, matchKeyField }) {
  const { data, update } = useStore();
  const targetRows = readPath(data, targetPath);

  useEffect(() => {
    const autoRows = computeAutoRows(data) || [];
    const existing = Array.isArray(targetRows) ? targetRows.filter(r => r != null) : [];

    const manualRows = existing.filter(r => !r?.[matchKeyField]);
    const freshAuto = autoRows.map(auto => {
      const prev = existing.find(r => r?.[matchKeyField] && r[matchKeyField] === auto[matchKeyField]);
      return prev ? { ...prev, ...auto } : auto;
    });

    const merged = [...freshAuto, ...manualRows];

    if (!arraysShallowEqual(merged, existing)) {
      update(targetPath, merged);
    }
    // We intentionally depend on `data` (full snapshot) so any source change
    // re-evaluates. update is stable. The hook itself dedupes via the equality
    // check above, so this does not loop.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);
}

function readPath(obj, path) {
  if (!obj) return undefined;
  return path.split('.').reduce((cur, k) => (cur == null ? cur : cur[k]), obj);
}

function arraysShallowEqual(a, b) {
  if (a === b) return true;
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (!objectShallowEqual(a[i], b[i])) return false;
  }
  return true;
}

function objectShallowEqual(a, b) {
  if (a === b) return true;
  if (!a || !b) return false;
  const aKeys = Object.keys(a);
  const bKeys = Object.keys(b);
  if (aKeys.length !== bKeys.length) return false;
  for (const k of aKeys) {
    if (!Object.prototype.hasOwnProperty.call(b, k)) return false;
    const av = a[k];
    const bv = b[k];
    if (av === bv) continue;
    if (Array.isArray(av) && Array.isArray(bv)) {
      if (av.length !== bv.length) return false;
      for (let i = 0; i < av.length; i++) {
        if (av[i] !== bv[i]) return false;
      }
      continue;
    }
    return false;
  }
  return true;
}
