// Stable per-row identifier used as the external key when records are sent
// to Salesforce. Every row in every table-style array carries a `_uid`.

export function genUid() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return 'uid-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
}

export function withUid(row) {
  if (row && typeof row === 'object' && !row._uid) {
    return { ...row, _uid: genUid() };
  }
  return row;
}

// Recursively walk loaded state and stamp `_uid` on any object element of
// any array. Used once at load time to migrate older saved data.
export function ensureUidsDeep(value) {
  if (Array.isArray(value)) {
    return value.map(item => {
      if (item && typeof item === 'object') {
        const stamped = item._uid ? item : { ...item, _uid: genUid() };
        return ensureUidsDeep(stamped);
      }
      return item;
    });
  }
  if (value && typeof value === 'object') {
    const next = {};
    for (const k of Object.keys(value)) next[k] = ensureUidsDeep(value[k]);
    return next;
  }
  return value;
}
