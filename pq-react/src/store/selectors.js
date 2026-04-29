// Pure selector helpers that read from the store data shape.

export function getOwnerNameOptions(data) {
  const family = data?.family || {};
  const flags = data?._flags || {};

  const c1 = [family.client1FirstName, family.client1LastName]
    .map(s => (s || '').trim())
    .filter(Boolean)
    .join(' ');
  const c2 = [family.client2FirstName, family.client2LastName]
    .map(s => (s || '').trim())
    .filter(Boolean)
    .join(' ');

  const opts = [];
  if (c1) opts.push(c1);
  if ((flags.hasSpouse || family.client2FirstName || family.client2LastName) && c2) {
    opts.push(c2);
  }
  if (c1 && c2) opts.push('Joint');
  opts.push('Trust');
  return opts;
}
