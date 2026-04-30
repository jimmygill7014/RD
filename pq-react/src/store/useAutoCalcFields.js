import { useEffect } from 'react';
import { useStore } from './StoreContext.jsx';

function calcAge(dob) {
  if (!dob) return '';
  const d = new Date(dob);
  if (isNaN(d.getTime())) return '';
  const today = new Date();
  let age = today.getFullYear() - d.getFullYear();
  const m = today.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < d.getDate())) age--;
  return age;
}

function calcAgeAtDate(dob, targetDate) {
  if (!dob || !targetDate) return '';
  const d = new Date(dob);
  const t = new Date(targetDate);
  if (isNaN(d.getTime()) || isNaN(t.getTime())) return '';
  let age = t.getFullYear() - d.getFullYear();
  const m = t.getMonth() - d.getMonth();
  if (m < 0 || (m === 0 && t.getDate() < d.getDate())) age--;
  return age;
}

function countRowsWithData(rows) {
  if (!Array.isArray(rows)) return 0;
  return rows.filter(r => {
    if (!r) return false;
    const name = String(r.name || '').trim();
    const age = String(r.age || '').trim();
    return name !== '' || age !== '';
  }).length;
}

/**
 * Reactively auto-calculates derived family/employment fields:
 *   - family.client1Age / client2Age  ← calcAge(DOB)
 *   - employment.client{1,2}.retirementAge ← calcAgeAtDate(DOB, retirementDate)
 *   - family.numChildren / numGrandchildren ← count of populated rows
 *
 * Writes back via update() only when the computed value differs from the
 * stored one. Mirrors the legacy recalcPQComputed behavior.
 */
export function useAutoCalcFields() {
  const { data, update } = useStore();

  const fam = data?.family || {};
  const emp = data?.employment || {};

  const c1Age = calcAge(fam.client1DOB);
  const c2Age = calcAge(fam.client2DOB);
  const c1RetAge = calcAgeAtDate(fam.client1DOB, emp?.client1?.retirementDate);
  const c2RetAge = calcAgeAtDate(fam.client2DOB, emp?.client2?.retirementDate);
  const childCount = countRowsWithData(fam.children);
  const grandCount = countRowsWithData(fam.grandchildren);

  useEffect(() => {
    if (c1Age !== '' && String(c1Age) !== String(fam.client1Age ?? '')) {
      update('family.client1Age', String(c1Age));
    }
    if (c2Age !== '' && String(c2Age) !== String(fam.client2Age ?? '')) {
      update('family.client2Age', String(c2Age));
    }
    if (c1RetAge !== '' && String(c1RetAge) !== String(emp?.client1?.retirementAge ?? '')) {
      update('employment.client1.retirementAge', String(c1RetAge));
    }
    if (c2RetAge !== '' && String(c2RetAge) !== String(emp?.client2?.retirementAge ?? '')) {
      update('employment.client2.retirementAge', String(c2RetAge));
    }
    // Child counts only override when computed > 0 (legacy: keep manual entry
    // when no rows have data yet).
    if (childCount > 0 && String(childCount) !== String(fam.numChildren ?? '')) {
      update('family.numChildren', String(childCount));
    }
    if (grandCount > 0 && String(grandCount) !== String(fam.numGrandchildren ?? '')) {
      update('family.numGrandchildren', String(grandCount));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [c1Age, c2Age, c1RetAge, c2RetAge, childCount, grandCount,
      fam.client1Age, fam.client2Age, fam.numChildren, fam.numGrandchildren,
      emp?.client1?.retirementAge, emp?.client2?.retirementAge]);
}
