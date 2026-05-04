import { formatDollars, getDashboardMetrics } from '../store/selectors.js';

const num = v => {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
};
const cleanRows = arr => (arr || []).filter(r =>
  r && Object.entries(r).some(([k, v]) => !k.startsWith('_') && v != null && String(v).trim() !== '')
);

function Row({ label, value, total }) {
  return (
    <div className={'plan-sum-row' + (total ? ' is-total' : '')}>
      <span className="plan-sum-label">{label}</span>
      <span className="plan-sum-value">{value}</span>
    </div>
  );
}

function AcctRow({ label, sub, value, indent }) {
  return (
    <div className={'plan-sum-row plan-sum-row-multi' + (indent ? ' plan-sum-row-indent' : '')}>
      <div className="plan-sum-row-main">
        <span className="plan-sum-label">{label}</span>
        {sub && <span className="plan-sum-sublabel">{sub}</span>}
      </div>
      {value && <span className="plan-sum-value">{value}</span>}
    </div>
  );
}

function GroupHead({ label, value }) {
  return (
    <div className="plan-sum-grouphead">
      <span className="plan-sum-grouphead-label">{label}</span>
      {value && <span className="plan-sum-grouphead-value">{value}</span>}
    </div>
  );
}

function Section({ title, children }) {
  return (
    <div className="plan-sum-section">
      <div className="plan-sum-title">{title}</div>
      {children}
    </div>
  );
}

function PersonEmpBlock({ name, status, e }) {
  const sub = [e.jobTitle, e.employer].filter(Boolean).join(' @ ');
  return (
    <>
      <Row label={name} value={status || '—'} total />
      {sub && <AcctRow label="Occupation" value={sub} indent />}
      {num(e.numOfYears) > 0 && <AcctRow label="# of Years" value={String(e.numOfYears)} indent />}
      {e.retirementDate && <AcctRow label="Retirement Date" value={e.retirementDate} indent />}
    </>
  );
}

export default function PlanSummaryPanel({ data }) {
  const m = getDashboardMetrics(data);

  const family   = data.family        || {};
  const contact  = data.contact       || {};
  const assets   = data.assets        || {};
  const liabs    = data.liabilities   || {};
  const income   = data.income        || {};
  const taxes    = data.taxesExpenses || {};
  const ins      = data.insurance     || {};
  const rels     = data.relationships || {};
  const emp      = data.employment    || {};
  const epData   = data.goals         || {};
  const goalsTxt = (data._goals || '').trim();
  const flags    = data._flags        || {};

  const c1Name = [family.client1FirstName, family.client1LastName].filter(Boolean).join(' ') || 'Client 1';
  const c2Name = [family.client2FirstName, family.client2LastName].filter(Boolean).join(' ');
  const clientDisplay = c2Name ? `${c1Name} & ${c2Name}` : c1Name;
  const age1 = family.client1Age ? ` (age ${family.client1Age})` : '';
  const age2 = family.client2Age ? ` (age ${family.client2Age})` : '';
  const ageDisplay = c2Name ? `${c1Name}${age1} & ${c2Name}${age2}` : `${c1Name}${age1}`;
  const cityState = [contact.city, contact.state].filter(Boolean).join(', ');

  const c1Status = emp.client1?.status ?? 'Employed';
  const c2Status = emp.client2?.status ?? 'Employed';

  const childrenRows = (family.children || []).filter(c => c && (c.name || c.age));
  const grandRows = (family.grandchildren || []).filter(c => c && (c.name || c.age));

  // Investment accounts to summarize
  const acctGroups = [
    { key: 'taxDeferred', label: 'Tax-Deferred', total: m.taxDeferred },
    { key: 'roth',        label: 'Roth',         total: m.taxFree - (data?.assets?.hsa ? acctSum(data.assets.hsa, 'marketValue') : 0) },
    { key: 'taxable',     label: 'Taxable',      total: acctSum(assets.taxable, 'marketValue') },
    { key: 'cashCd',      label: 'Cash / CD',    total: acctSum(assets.cashCd, 'marketValue') },
    { key: 'plan529',     label: '529 Plan',     total: acctSum(assets.plan529, 'marketValue') },
    { key: 'hsa',         label: 'HSA',          total: acctSum(assets.hsa, 'marketValue') },
  ];
  const hasAnyAcct = acctGroups.some(g => cleanRows(assets[g.key]).length > 0);

  const reRows = cleanRows(assets.realEstate);
  const busRows = cleanRows(assets.businessOther);
  const liabRows = cleanRows(liabs.items);

  const incomeGroups = [
    { key: 'employment',     label: 'Employment' },
    { key: 'socialSecurity', label: 'Social Security' },
    { key: 'pension',        label: 'Pension' },
    { key: 'other',          label: 'Other' },
  ];
  const hasAnyInc = incomeGroups.some(g => cleanRows(income[g.key]).length > 0);

  const expRows = cleanRows(taxes.expenses);
  const insRows = cleanRows(ins.policies);
  const relRows = cleanRows(rels.professionals);

  return (
    <>
      {/* Header */}
      <div className="plan-sum-header">
        <div className="plan-sum-client">{clientDisplay}</div>
        {ageDisplay !== clientDisplay && <div className="plan-sum-ages">{ageDisplay}</div>}
        {cityState && <div className="plan-sum-location">{cityState}</div>}
      </div>

      {/* Contact */}
      {(contact.client1Email || contact.homePhone || contact.address1 || contact.city) && (
        <Section title="Contact">
          {contact.client1Email && <Row label="Email" value={contact.client1Email} />}
          {contact.homePhone && <Row label="Phone" value={contact.homePhone} />}
          {[contact.address1, contact.city, contact.state, contact.zip].filter(Boolean).length > 0 && (
            <Row label="Address" value={[contact.address1, contact.city, contact.state, contact.zip].filter(Boolean).join(', ')} />
          )}
        </Section>
      )}

      {/* Family */}
      {(childrenRows.length > 0 || grandRows.length > 0) && (
        <Section title="Family">
          {childrenRows.length > 0 && <>
            <Row label="Children" value={String(childrenRows.length)} total />
            {childrenRows.map((c, i) => (
              <AcctRow
                key={i}
                label={c.name || 'Child'}
                sub={c.age ? `age ${c.age}` : ''}
                value=""
                indent
              />
            ))}
          </>}
          {grandRows.length > 0 && <>
            <Row label="Grandchildren" value={String(grandRows.length)} total />
            {grandRows.map((c, i) => (
              <AcctRow
                key={i}
                label={c.name || 'Grandchild'}
                sub={c.age ? `age ${c.age}` : ''}
                value=""
                indent
              />
            ))}
          </>}
        </Section>
      )}

      {/* Employment */}
      <Section title="Employment">
        <PersonEmpBlock name={c1Name} status={c1Status} e={emp.client1 || {}} />
        {flags.hasSpouse && (
          <PersonEmpBlock name={c2Name || 'Spouse'} status={c2Status} e={emp.client2 || {}} />
        )}
      </Section>

      {/* Net Worth */}
      <Section title="Net Worth">
        <div className={'plan-sum-big ' + (m.totalNetWorth >= 0 ? 'plan-sum-positive' : 'plan-sum-negative')}>
          {formatDollars(m.totalNetWorth)}
        </div>
        <Row label="Total Assets" value={formatDollars(m.totalAssets)} />
        <Row label="Total Liabilities" value={formatDollars(m.totalLiabilities)} total />
      </Section>

      {/* Investment Accounts */}
      {hasAnyAcct && (
        <Section title="Investment Accounts">
          {acctGroups.map(g => {
            const rows = cleanRows(assets[g.key]);
            if (!rows.length) return null;
            return (
              <div key={g.key}>
                <GroupHead label={g.label} value={formatDollars(g.total)} />
                {rows.map((r, i) => {
                  const sub = [
                    r.type, r.ownership,
                    r.beneficiary ? 'Bene: ' + r.beneficiary : '',
                    num(r.personalAdditions) > 0 ? 'Adds: ' + formatDollars(num(r.personalAdditions)) + '/yr' : '',
                  ].filter(Boolean).join(' · ');
                  return (
                    <AcctRow
                      key={i}
                      label={r.custodian || r.description || 'Account'}
                      sub={sub}
                      value={formatDollars(num(r.marketValue))}
                      indent
                    />
                  );
                })}
              </div>
            );
          })}
          <Row label="Investable Total" value={formatDollars(m.investableAssets)} total />
        </Section>
      )}

      {/* Real Estate */}
      {reRows.length > 0 && (
        <Section title="Real Estate">
          {reRows.map((r, i) => {
            const subParts = [r.ownership];
            if (num(r.remainingLoan)) subParts.push('Loan: ' + formatDollars(num(r.remainingLoan)));
            if (num(r.payment))       subParts.push('Pmt: ' + formatDollars(num(r.payment)) + '/mo');
            if (r.interestRate)       subParts.push(r.interestRate + '%');
            return (
              <AcctRow
                key={i}
                label={r.description || 'Property'}
                sub={subParts.filter(Boolean).join(' · ')}
                value={formatDollars(num(r.marketValue))}
              />
            );
          })}
          <Row label="Total" value={formatDollars(m.realEstateAssets)} total />
        </Section>
      )}

      {/* Business & Other */}
      {busRows.length > 0 && (
        <Section title="Business & Other Assets">
          {busRows.map((r, i) => {
            const subParts = [r.ownership, num(r.costBasis) ? 'Basis: ' + formatDollars(num(r.costBasis)) : ''].filter(Boolean);
            return (
              <AcctRow
                key={i}
                label={r.description || 'Asset'}
                sub={subParts.join(' · ')}
                value={formatDollars(num(r.marketValue))}
              />
            );
          })}
          <Row label="Total" value={formatDollars(m.businessOtherAssets)} total />
        </Section>
      )}

      {/* Liabilities */}
      {liabRows.length > 0 && (
        <Section title="Liabilities">
          {liabRows.map((r, i) => {
            const subParts = [];
            if (num(r.payment)) subParts.push('Pmt: ' + formatDollars(num(r.payment)) + '/mo');
            if (r.interestRate) subParts.push(r.interestRate + '%');
            if (r.term) subParts.push(r.term);
            return (
              <AcctRow
                key={i}
                label={r.description || 'Liability'}
                sub={subParts.join(' · ')}
                value={formatDollars(num(r.amount))}
              />
            );
          })}
          <Row label="Total Liabilities" value={formatDollars(m.totalLiabilities)} total />
        </Section>
      )}

      {/* Income Sources */}
      {hasAnyInc && (
        <Section title="Income Sources">
          {incomeGroups.map(g => {
            const rows = cleanRows(income[g.key]);
            if (!rows.length) return null;
            const total = rows.reduce((sum, r) => sum + num(r.annualAmount), 0);
            return (
              <div key={g.key}>
                <GroupHead label={g.label} value={formatDollars(total)} />
                {rows.map((r, i) => {
                  const subParts = [];
                  if (r.owner && r.description) subParts.push(r.owner);
                  if (r.startDate) subParts.push('Start: ' + r.startDate);
                  if (r.cola) subParts.push('COLA: ' + r.cola + '%');
                  return (
                    <AcctRow
                      key={i}
                      label={r.description || r.owner || g.label}
                      sub={subParts.join(' · ')}
                      value={formatDollars(num(r.annualAmount))}
                      indent
                    />
                  );
                })}
              </div>
            );
          })}
          <Row label="Total Income" value={formatDollars(m.totalIncome)} total />
        </Section>
      )}

      {/* Annual Cash Flow */}
      <Section title="Annual Cash Flow">
        <Row label="Gross Income" value={formatDollars(m.totalIncome)} />
        {m.totalTax > 0           && <Row label="Less: Tax"        value={'(' + formatDollars(m.totalTax) + ')'} />}
        {m.totalSavings > 0       && <Row label="Less: Savings"    value={'(' + formatDollars(m.totalSavings) + ')'} />}
        {m.liabilityPayments > 0  && <Row label="Less: Debt Pmts"  value={'(' + formatDollars(m.liabilityPayments) + ')'} />}
        {m.nonLiabilityExpenses > 0 && <Row label="Less: Living Exp" value={'(' + formatDollars(m.nonLiabilityExpenses) + ')'} />}
        <div className={'plan-sum-row is-total'}>
          <span className="plan-sum-label">Net Cash Flow</span>
          <span className={'plan-sum-value ' + (m.cashFlow >= 0 ? 'plan-sum-positive' : 'plan-sum-negative')}>
            {formatDollars(m.cashFlow)}
          </span>
        </div>
      </Section>

      {/* Taxes */}
      {m.totalTax > 0 && (
        <Section title="Taxes">
          {num(taxes.federalTax) > 0 && <Row label="Federal" value={formatDollars(num(taxes.federalTax))} />}
          {num(taxes.stateTax) > 0   && <Row label="State"   value={formatDollars(num(taxes.stateTax))} />}
          {num(taxes.ficaTax) > 0    && <Row label="FICA"    value={formatDollars(num(taxes.ficaTax))} />}
          <Row label="Total Tax" value={formatDollars(m.totalTax)} total />
          {m.totalIncome > 0 && (
            <Row label="Effective Rate" value={((m.totalTax / m.totalIncome) * 100).toFixed(1) + '%'} />
          )}
        </Section>
      )}

      {/* Annual Expenses */}
      {expRows.length > 0 && (
        <Section title="Annual Expenses">
          {expRows.map((r, i) => (
            <AcctRow
              key={i}
              label={r.description || 'Expense'}
              sub={r.notes || ''}
              value={formatDollars(num(r.amount))}
            />
          ))}
          <Row
            label="Total Expenses"
            value={formatDollars(expRows.reduce((s, r) => s + num(r.amount), 0))}
            total
          />
        </Section>
      )}

      {/* Insurance */}
      {insRows.length > 0 && (
        <Section title="Insurance">
          {insRows.map((r, i) => {
            const label = [r.company, r.type].filter(Boolean).join(' — ') || 'Policy';
            const subParts = [];
            if (r.insured)        subParts.push('Insured: ' + r.insured);
            if (r.owner)          subParts.push('Owner: ' + r.owner);
            if (num(r.benefit))   subParts.push('Benefit: ' + formatDollars(num(r.benefit)));
            if (num(r.cashValue)) subParts.push('CV: ' + formatDollars(num(r.cashValue)));
            if (r.beneficiary)    subParts.push('Bene: ' + r.beneficiary);
            const val = num(r.annualPremium) ? formatDollars(num(r.annualPremium)) + '/yr' : '';
            return (
              <AcctRow
                key={i}
                label={label}
                sub={subParts.join(' · ')}
                value={val}
              />
            );
          })}
        </Section>
      )}

      {/* Professional Relationships */}
      {relRows.length > 0 && (
        <Section title="Professional Relationships">
          {relRows.map((r, i) => {
            const subParts = [r.firm, [r.city, r.state].filter(Boolean).join(', ')].filter(Boolean);
            return (
              <AcctRow
                key={i}
                label={r.role || 'Professional'}
                sub={subParts.join(' · ')}
                value={r.name || '—'}
              />
            );
          })}
        </Section>
      )}

      {/* Estate Plan & Goals */}
      {((Array.isArray(epData.estatePlan) && epData.estatePlan.length) || epData.yearsEstablished || goalsTxt) && (
        <Section title="Estate Plan & Goals">
          {Array.isArray(epData.estatePlan) && epData.estatePlan.length > 0 && (
            <Row label="Plan Components" value={epData.estatePlan.join(', ')} />
          )}
          {epData.yearsEstablished && <Row label="Year Established" value={epData.yearsEstablished} />}
          {goalsTxt && <div className="plan-sum-notes">{goalsTxt}</div>}
        </Section>
      )}
    </>
  );
}

function acctSum(rows, key) {
  if (!Array.isArray(rows)) return 0;
  return rows.reduce((s, r) => s + (r ? num(r[key]) : 0), 0);
}
