import { useStore } from '../store/StoreContext.jsx';
import { formatDollars, getDashboardMetrics } from '../store/selectors.js';
import TaxTriangleSVG from '../components/TaxTriangleSVG.jsx';

const num = v => {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
};
const sumA = (arr, key) => (arr || []).filter(Boolean).reduce((s, r) => s + num(r[key]), 0);
const cleanRows = arr => (arr || []).filter(r =>
  r && Object.entries(r).some(([k, v]) => !k.startsWith('_') && v != null && String(v).trim() !== '')
);
const dash = v => (v == null || v === '') ? '—' : v;

function PresSection({ title, badge, compact, children }) {
  return (
    <div className={'pres-section' + (compact ? ' is-compact' : '')}>
      <div className="pres-section-header">
        <h2 className="pres-section-title">{title}</h2>
        {badge != null && <span className="pres-section-badge">{badge}</span>}
      </div>
      <div className="pres-section-body">{children}</div>
    </div>
  );
}

function PresTable({ columns, rows, footTotals }) {
  return (
    <div className="pres-table-wrap">
      <table className="pres-table">
        <thead>
          <tr>
            {columns.map(c => (
              <th key={c.key} className={c.right ? 'th-right' : ''}>{c.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} className={i % 2 !== 0 ? 'pres-row-alt' : ''}>
              {columns.map(c => {
                const v = row[c.key];
                const display = v != null && v !== '' ? v : '—';
                return (
                  <td key={c.key} className={c.right ? 'td-right td-mono' : ''}>
                    {display}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
        {footTotals && (
          <tfoot>
            <tr>
              {columns.map((c, i) => {
                const total = footTotals[c.key];
                if (total != null) {
                  return <td key={c.key} className={c.right ? 'td-right td-mono' : ''}>{total}</td>;
                }
                if (i === 0) return <td key={c.key} style={{ fontWeight: 600 }}>Total</td>;
                return <td key={c.key} />;
              })}
            </tr>
          </tfoot>
        )}
      </table>
    </div>
  );
}

function InfoGrid({ items }) {
  return (
    <div className="pres-info-grid">
      {items.filter(it => !it.hidden).map((it, i) => (
        <div key={i} className={'pres-info-cell' + (it.span ? ` span-${it.span}` : '')}>
          <div className="pres-info-label">{it.label}</div>
          <div className="pres-info-value">{dash(it.value)}</div>
        </div>
      ))}
    </div>
  );
}

function presRow(label, value) {
  return (
    <div className="pres-row" style={{ display: 'flex', justifyContent: 'space-between' }}>
      <span>{label}</span>
      <span>{value}</span>
    </div>
  );
}

export default function PresentationScreen({ onBack }) {
  const { data } = useStore();
  const m = getDashboardMetrics(data);

  const family = data.family || {};
  const contact = data.contact || {};
  const emp = data.employment || {};
  const assets = data.assets || {};
  const liabs = data.liabilities || {};
  const ins = data.insurance || {};
  const income = data.income || {};
  const taxes = data.taxesExpenses || {};
  const rels = data.relationships || {};
  const goalsText = (data._goals || '').trim();
  const epData = data.goals || {};

  const c1First = (family.client1FirstName || '').trim();
  const c1Last  = (family.client1LastName || '').trim();
  const c2First = (family.client2FirstName || '').trim();
  const c2Last  = (family.client2LastName || '').trim();
  const c1Name = [c1First, c1Last].filter(Boolean).join(' ') || 'Client';
  const c2Name = [c2First, c2Last].filter(Boolean).join(' ');
  const clientDisplay = c2Name ? `${c1Name} & ${c2Name}` : c1Name;
  const age1 = family.client1Age ? ` (age ${family.client1Age})` : '';
  const age2 = family.client2Age ? ` (age ${family.client2Age})` : '';
  const clientAges = c2Name ? `${c1Name}${age1} & ${c2Name}${age2}` : `${c1Name}${age1}`;
  const cityState = [contact.city, contact.state].filter(Boolean).join(', ');

  const reAssets   = sumA(assets.realEstate,    'marketValue');
  const busOther   = sumA(assets.businessOther, 'marketValue');
  const taxDef     = sumA(assets.taxDeferred,   'marketValue');
  const rothVal    = sumA(assets.roth,          'marketValue');
  const taxableVal = sumA(assets.taxable,       'marketValue');
  const cashCd     = sumA(assets.cashCd,        'marketValue');
  const plan529    = sumA(assets.plan529,       'marketValue');
  const hsaVal     = sumA(assets.hsa,           'marketValue');

  const empInc   = sumA(income.employment,     'annualAmount');
  const ssInc    = sumA(income.socialSecurity, 'annualAmount');
  const pensInc  = sumA(income.pension,        'annualAmount');
  const otherInc = sumA(income.other,          'annualAmount');

  const fedTax   = num(taxes.federalTax);
  const stTax    = num(taxes.stateTax);
  const ficaTax  = num(taxes.ficaTax);
  const totalTax = fedTax + stTax + ficaTax;

  const livingExpField = num(taxes.livingExpenses);
  const expRows = (taxes.expenses || []).filter(Boolean);
  const debtPayments = (liabs.items || []).filter(Boolean).reduce((s, r) => s + num(r.payment) * 12, 0);

  const totalSavings = sumA(assets.taxDeferred, 'personalAdditions')
    + sumA(assets.roth, 'personalAdditions')
    + sumA(assets.taxable, 'personalAdditions');

  const totalIncome = empInc + ssInc + pensInc + otherInc; // sum-all (not the time-filtered dashboard total)
  const totalAssets = reAssets + busOther + taxDef + rothVal + taxableVal + cashCd + plan529 + hsaVal;
  const totalLiab = sumA(liabs.items, 'amount');
  const netWorth = totalAssets - totalLiab;
  const livingExpAll = livingExpField + expRows.filter(r => r._source !== 'liability' && r._source !== 'insurance')
    .reduce((s, r) => s + num(r.amount), 0);
  const netCF = totalIncome - totalTax - totalSavings - debtPayments - livingExpAll;
  const savingsRatio = totalIncome > 0 ? (totalSavings / totalIncome) * 100 : 0;
  const dtiRatio     = totalIncome > 0 ? (debtPayments / totalIncome) * 100 : 0;

  const childRows = (family.children || []).filter(c => c && (c.name || c.age));
  const grandRows = (family.grandchildren || []).filter(c => c && (c.name || c.age));

  // Income tax tri-mix
  const taxFreeTotal = rothVal + hsaVal;
  const taxableTotal = taxableVal + cashCd;
  const taxDeferredTotal = taxDef;
  const triTotal = taxFreeTotal + taxableTotal + taxDeferredTotal;

  return (
    <section className="screen is-active">
      <div className="pres-nav no-print">
        <div className="pres-nav-brand brand" aria-label="Pure Financial Advisors">
          <span className="brand-text">Pure Financial Advisors</span>
        </div>
        <div className="pres-nav-actions">
          <button type="button" className="btn btn-light" onClick={onBack}>← Edit</button>
          <button type="button" className="btn btn-primary" onClick={() => window.print()}>
            Print / Save PDF
          </button>
        </div>
      </div>

      <div className="pres-content">
        {/* Header */}
        <div className="pres-client-header">
          <div>
            <div className="pres-client-name">{clientDisplay}</div>
            <div className="pres-client-sub">
              {clientAges}{cityState ? '  ·  ' + cityState : ''}
            </div>
          </div>
          <div className="pres-prepared">
            <strong>Pure Financial Advisors</strong>
            Prepared: {new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}
          </div>
        </div>

        {/* KPI cards */}
        <div className="pres-kpi-row">
          {[
            { label: 'Net Worth',         value: formatDollars(netWorth),    color: '#1f3f67', cls: netWorth >= 0 ? 'positive' : 'negative' },
            { label: 'Total Assets',      value: formatDollars(totalAssets), color: '#4f46e5', cls: '' },
            { label: 'Total Liabilities', value: formatDollars(totalLiab),   color: '#e11d48', cls: '' },
            { label: 'Annual Income',     value: formatDollars(totalIncome), color: '#059669', cls: '' },
            { label: 'Net Cash Flow',     value: formatDollars(netCF),       color: netCF >= 0 ? '#059669' : '#e11d48', cls: netCF >= 0 ? 'positive' : 'negative' },
          ].map(k => (
            <div key={k.label} className="pres-kpi-card" style={{ '--kpi-color': k.color }}>
              <div className="pres-kpi-label">{k.label}</div>
              <div className={'pres-kpi-value ' + k.cls}>{k.value}</div>
            </div>
          ))}
        </div>

        {/* Family */}
        <PresSection title="Family Information">
          <FamilyBlock
            family={family}
            contact={contact}
            c1First={c1First}
            c1Last={c1Last}
            c2First={c2First}
            c2Last={c2Last}
            childRows={childRows}
            grandRows={grandRows}
          />
        </PresSection>

        {/* Occupation */}
        {(emp.client1?.status || emp.client2?.status || emp.client1?.jobTitle || emp.client2?.jobTitle || emp.client1?.employer || emp.client2?.employer || emp.client1?.lastJob || emp.client2?.lastJob) && (
          <PresSection title="Occupation" compact>
            <EmpBlock who={c1Name} e={emp.client1 || {}} />
            {(c2Name || emp.client2?.status) && (
              <EmpBlock who={c2Name || 'Spouse'} e={emp.client2 || {}} />
            )}
          </PresSection>
        )}

        {/* Professional Relationships */}
        {cleanRows(rels.professionals).length > 0 && (
          <PresSection title="Current Professional Relationships" compact>
            <PresTable
              columns={[
                { label: 'Role',  key: 'role' },
                { label: 'Name',  key: 'name' },
                { label: 'Firm',  key: 'firm' },
                { label: 'City',  key: 'city' },
                { label: 'State', key: 'state' },
              ]}
              rows={cleanRows(rels.professionals)}
            />
          </PresSection>
        )}

        {/* Goals */}
        {goalsText && (
          <PresSection title="Goals" compact>
            <p className="pres-goals-text">{goalsText}</p>
          </PresSection>
        )}

        {/* Real Estate */}
        {cleanRows(assets.realEstate).length > 0 && (
          <RealEstateSection rows={cleanRows(assets.realEstate)} reAssets={reAssets} />
        )}

        {/* Investments + Tax Triangle */}
        {(cleanRows(assets.taxDeferred).length || cleanRows(assets.roth).length || cleanRows(assets.taxable).length) > 0 && (
          <InvestmentSection
            tdRows={cleanRows(assets.taxDeferred)}
            rothRows={cleanRows(assets.roth)}
            txRows={cleanRows(assets.taxable)}
            taxDef={taxDef}
            rothVal={rothVal}
            taxableVal={taxableVal}
            triData={triTotal ? { taxFree: taxFreeTotal, taxable: taxableTotal, taxDeferred: taxDeferredTotal } : null}
          />
        )}

        {/* Cash, CDs, Business & Other */}
        {(cleanRows(assets.cashCd).length || cleanRows(assets.plan529).length || cleanRows(assets.hsa).length || cleanRows(assets.businessOther).length) > 0 && (
          <CashBusinessSection
            cashRows={cleanRows(assets.cashCd)}
            plan529Rows={cleanRows(assets.plan529)}
            hsaRows={cleanRows(assets.hsa)}
            busRows={cleanRows(assets.businessOther)}
            cashCd={cashCd}
            plan529={plan529}
            hsaVal={hsaVal}
            busOther={busOther}
          />
        )}

        {/* Liabilities + Net Worth */}
        <LiabilitiesSection
          rows={cleanRows(liabs.items)}
          totalLiab={totalLiab}
          totalAssets={totalAssets}
          netWorth={netWorth}
        />

        {/* Insurance */}
        {cleanRows(ins.policies).length > 0 && (
          <InsuranceSection rows={cleanRows(ins.policies)} policies={ins.policies} />
        )}

        {/* Income & Tax */}
        {((cleanRows(income.employment).length + cleanRows(income.socialSecurity).length + cleanRows(income.pension).length + cleanRows(income.other).length) > 0
          || num(taxes.capLossCarryForward) || num(taxes.taxableIncome) || num(taxes.standardItemDeduction) || totalTax) > 0 && (
          <IncomeTaxSection
            empRows={cleanRows(income.employment)}
            ssRows={cleanRows(income.socialSecurity)}
            pensRows={cleanRows(income.pension)}
            otherRows={cleanRows(income.other)}
            empInc={empInc} ssInc={ssInc} pensInc={pensInc} otherInc={otherInc}
            totalIncome={totalIncome}
            taxes={taxes} fedTax={fedTax} stTax={stTax} ficaTax={ficaTax} totalTax={totalTax}
          />
        )}

        {/* Expenses & Cash Flow */}
        {(livingExpField || expRows.length || totalIncome) > 0 && (
          <ExpensesCashFlowSection
            livingExpField={livingExpField}
            expRows={expRows}
            totalIncome={totalIncome}
            totalTax={totalTax}
            ficaTax={ficaTax}
            totalSavings={totalSavings}
            debtPayments={debtPayments}
            livingExp={livingExpAll}
            netCF={netCF}
            savingsRatio={savingsRatio}
            dtiRatio={dtiRatio}
          />
        )}

        {/* Estate Plan */}
        {((Array.isArray(epData.estatePlan) && epData.estatePlan.length) || epData.yearsEstablished) && (
          <PresSection title="Estate Plan">
            <InfoGrid items={[
              { label: 'Plan Components', value: Array.isArray(epData.estatePlan) ? epData.estatePlan.join(', ') : epData.estatePlan, span: 2 },
              { label: 'Year Established / Updated', value: epData.yearsEstablished },
            ]} />
          </PresSection>
        )}
      </div>
    </section>
  );
}

/* ---------- Section sub-components ---------- */

function FamilyBlock({ family, contact, c1First, c1Last, c2First, c2Last, childRows, grandRows }) {
  const hasSpouseRow = c2First || c2Last || family.client2Nickname || family.client2DOB || family.client2Age;
  const hasAddrRow = contact.address1 || contact.address2 || contact.city || contact.state || contact.zip;

  const numChildren = parseInt(family.numChildren, 10) || 0;
  const numGrand = parseInt(family.numGrandchildren, 10) || 0;
  const showChildren = childRows.length > 0 || numChildren > 0;
  const showGrand = grandRows.length > 0 || numGrand > 0;
  const hasKids = showChildren || showGrand;

  const Cell = ({ label, value, span, empty }) => (
    <div className={'pres-ss-cell' + (span ? ` span-${span}` : '') + (empty ? ' is-empty' : '')}>
      {!empty && (
        <>
          <div className="pres-ss-label">{label}</div>
          <div className="pres-ss-value">{dash(value)}</div>
        </>
      )}
    </div>
  );

  const ssGrid = (
    <div className="pres-ss-grid pres-ss-family">
      <Cell label="Your Name"  value={c1First} />
      <Cell label="Last Name"  value={c1Last} />
      <Cell label="Nickname"   value={family.client1Nickname} />
      <Cell label="Birth Date" value={family.client1DOB} />
      <Cell label="Age"        value={family.client1Age} />

      {hasSpouseRow && <>
        <Cell label="Spouse's Name" value={c2First} />
        <Cell label="Last Name"     value={c2Last} />
        <Cell label="Nickname"      value={family.client2Nickname} />
        <Cell label="Birth Date"    value={family.client2DOB} />
        <Cell label="Age"           value={family.client2Age} />
      </>}

      {hasAddrRow && <>
        <Cell label="Residence Address" value={[contact.address1, contact.address2].filter(Boolean).join(', ')} span={2} />
        <Cell label="City"     value={contact.city} />
        <Cell label="State"    value={contact.state} />
        <Cell label="Zip Code" value={contact.zip} />
      </>}

      {contact.referredBy && <Cell label="Referred By" value={contact.referredBy} span={5} />}
    </div>
  );

  const buildKidGroup = (title, count, rows) => (
    <div className="pres-kids-group">
      <div className="pres-kids-heading">
        <span className="pres-kids-title">{title}</span>
        {count > 0 && <span className="pres-kids-count">{String(count)}</span>}
      </div>
      {rows.length ? (
        <ul className="pres-kids-list">
          {rows.map((r, i) => (
            <li key={i} className="pres-kids-item">
              <span className="pres-kids-name">{dash(r.name)}</span>
              <span className="pres-kids-age">{(r.age || r.age === 0) ? String(r.age) : '—'}</span>
            </li>
          ))}
        </ul>
      ) : (
        <div className="pres-kids-empty">No names recorded</div>
      )}
    </div>
  );

  return (
    <>
      {hasKids ? (
        <div className="pres-family-layout">
          {ssGrid}
          <div className={'pres-kids-panel ' + (showChildren && showGrand ? 'pres-kids-panel--split' : 'pres-kids-panel--single')}>
            {showChildren && buildKidGroup('Children', numChildren || childRows.length, childRows)}
            {showGrand && buildKidGroup('Grandchildren', numGrand || grandRows.length, grandRows)}
          </div>
        </div>
      ) : ssGrid}
      {family.additionalFamilyInfo && <>
        <div className="pres-subsec-label">Additional Family Information</div>
        <p className="pres-notes-text">{family.additionalFamilyInfo}</p>
      </>}
    </>
  );
}

function EmpBlock({ who, e }) {
  const status = e.status ?? 'Employed';
  const isWorking = status === 'Employed';
  return (
    <>
      <div className="pres-subsec-label">{who} · {status}</div>
      <InfoGrid items={[
        { label: 'Job Title',                              value: isWorking ? e.jobTitle : e.lastJob },
        { label: 'Employer' + (isWorking ? '' : ' (Last)'), value: isWorking ? e.employer : e.lastEmployer },
        { label: '# of Years',                             value: isWorking ? e.numOfYears : e.yearsWorked },
        { label: 'Retirement Date',                        value: e.retirementDate, hidden: !isWorking },
        { label: 'Retirement Age',                         value: e.retirementAge,  hidden: !isWorking },
        { label: 'Type of Business',                       value: e.typeOfBusiness, hidden: !isWorking },
      ]} />
    </>
  );
}

function RealEstateSection({ rows, reAssets }) {
  const reLoanTot = sumA(rows, 'remainingLoan');
  const rePmtTot  = sumA(rows, 'payment');
  const reEbtTot  = sumA(rows, 'incomeEBT');
  const reCbTot   = sumA(rows, 'costBasis');
  return (
    <PresSection title="Real Estate Assets" badge={formatDollars(reAssets)}>
      <PresTable
        columns={[
          { label: 'Description',  key: 'description' },
          { label: 'Market Value', key: '_mv',   right: true },
          { label: 'Loan Amount',  key: '_loan', right: true },
          { label: 'Int. Rate',    key: '_rate', right: true },
          { label: 'Term',         key: 'term',  right: true },
          { label: 'Pmt (P&I)',    key: '_pmt',  right: true },
          { label: 'Income EBT',   key: '_ebt',  right: true },
          { label: 'Year Acq.',    key: 'yearAcquired' },
          { label: 'Cost Basis',   key: '_cb',   right: true },
          { label: 'Ownership',    key: 'ownership' },
        ]}
        rows={rows.map(r => ({
          ...r,
          _mv:   formatDollars(num(r.marketValue)),
          _loan: num(r.remainingLoan) ? formatDollars(num(r.remainingLoan)) : '—',
          _rate: r.interestRate ? r.interestRate + '%' : '—',
          _pmt:  num(r.payment) ? formatDollars(num(r.payment)) : '—',
          _ebt:  num(r.incomeEBT) ? formatDollars(num(r.incomeEBT)) : '—',
          _cb:   num(r.costBasis) ? formatDollars(num(r.costBasis)) : '—',
        }))}
        footTotals={{
          _mv:   formatDollars(reAssets),
          _loan: reLoanTot ? formatDollars(reLoanTot) : '—',
          _pmt:  rePmtTot  ? formatDollars(rePmtTot)  : '—',
          _ebt:  reEbtTot  ? formatDollars(reEbtTot)  : '—',
          _cb:   reCbTot   ? formatDollars(reCbTot)   : '—',
        }}
      />
    </PresSection>
  );
}

function InvestmentSection({ tdRows, rothRows, txRows, taxDef, rothVal, taxableVal, triData }) {
  const invTotal = taxDef + rothVal + taxableVal;

  const GroupTitle = ({ label, amount, variant }) => (
    <div className={`pres-invest-group-title ${variant}`}>
      <span className="pres-invest-group-label">{label}</span>
      <span className="pres-invest-group-amount">{formatDollars(amount)}</span>
    </div>
  );

  return (
    <PresSection title="Investment Accounts" badge={formatDollars(invTotal)}>
      <div className="pres-invest-layout">
        <div className="pres-invest-accounts">
          {tdRows.length > 0 && <>
            <GroupTitle label="Tax-Deferred Retirement" amount={taxDef} variant="td" />
            <PresTable
              columns={[
                { label: 'Custodian',          key: 'custodian' },
                { label: 'Market Value',       key: '_mv',    right: true },
                { label: 'Personal Additions', key: '_add',   right: true },
                { label: 'Company Match',      key: '_match', right: true },
                { label: 'Type',               key: 'type' },
                { label: 'Ownership',          key: 'ownership' },
                { label: 'Beneficiary',        key: 'beneficiary' },
              ]}
              rows={tdRows.map(r => ({
                ...r,
                _mv:    formatDollars(num(r.marketValue)),
                _add:   num(r.personalAdditions) ? formatDollars(num(r.personalAdditions)) : '—',
                _match: num(r.companyMatch) ? formatDollars(num(r.companyMatch)) : '—',
              }))}
              footTotals={{
                _mv:    formatDollars(taxDef),
                _add:   sumA(tdRows, 'personalAdditions') ? formatDollars(sumA(tdRows, 'personalAdditions')) : '—',
                _match: sumA(tdRows, 'companyMatch') ? formatDollars(sumA(tdRows, 'companyMatch')) : '—',
              }}
            />
          </>}

          {rothRows.length > 0 && <>
            <GroupTitle label="Tax-Free Roth" amount={rothVal} variant="roth" />
            <PresTable
              columns={[
                { label: 'Custodian',          key: 'custodian' },
                { label: 'Market Value',       key: '_mv',    right: true },
                { label: 'Personal Additions', key: '_add',   right: true },
                { label: 'Company Match',      key: '_match', right: true },
                { label: 'Roth Type',          key: 'type' },
                { label: 'Ownership',          key: 'ownership' },
                { label: 'Beneficiary',        key: 'beneficiary' },
              ]}
              rows={rothRows.map(r => ({
                ...r,
                _mv:    formatDollars(num(r.marketValue)),
                _add:   num(r.personalAdditions) ? formatDollars(num(r.personalAdditions)) : '—',
                _match: num(r.companyMatch) ? formatDollars(num(r.companyMatch)) : '—',
              }))}
              footTotals={{
                _mv:    formatDollars(rothVal),
                _add:   sumA(rothRows, 'personalAdditions') ? formatDollars(sumA(rothRows, 'personalAdditions')) : '—',
                _match: sumA(rothRows, 'companyMatch') ? formatDollars(sumA(rothRows, 'companyMatch')) : '—',
              }}
            />
          </>}

          {txRows.length > 0 && <>
            <GroupTitle label="Taxable Non-Retirement" amount={taxableVal} variant="taxable" />
            <PresTable
              columns={[
                { label: 'Custodian',    key: 'custodian' },
                { label: 'Market Value', key: '_mv',  right: true },
                { label: 'Additions',    key: '_add', right: true },
                { label: 'Cost Basis',   key: '_cb',  right: true },
                { label: 'Ownership',    key: 'ownership' },
                { label: 'Var/Fixed',    key: 'variableFixed' },
                { label: 'Issue Date',   key: 'issueDate' },
                { label: 'Beneficiary',  key: 'beneficiary' },
              ]}
              rows={txRows.map(r => ({
                ...r,
                _mv:  formatDollars(num(r.marketValue)),
                _add: num(r.personalAdditions) ? formatDollars(num(r.personalAdditions)) : '—',
                _cb:  num(r.costBasis) ? formatDollars(num(r.costBasis)) : '—',
              }))}
              footTotals={{
                _mv:  formatDollars(taxableVal),
                _add: sumA(txRows, 'personalAdditions') ? formatDollars(sumA(txRows, 'personalAdditions')) : '—',
                _cb:  sumA(txRows, 'costBasis') ? formatDollars(sumA(txRows, 'costBasis')) : '—',
              }}
            />
          </>}
        </div>

        {triData && (
          <div className="pres-invest-triangle">
            <div className="pres-invest-tri-title">Tax Treatment Mix</div>
            <div className="pres-tri-chart">
              <TaxTriangleSVG
                taxFree={triData.taxFree}
                taxable={triData.taxable}
                taxDeferred={triData.taxDeferred}
                idSuffix="Pres"
              />
            </div>
            <div className="pres-invest-tri-footnote">Includes HSA &amp; cash in their tax categories.</div>
          </div>
        )}
      </div>
    </PresSection>
  );
}

function CashBusinessSection({ cashRows, plan529Rows, hsaRows, busRows, cashCd, plan529, hsaVal, busOther }) {
  const total = cashCd + plan529 + hsaVal + busOther;
  const hasLeft = cashRows.length || plan529Rows.length || hsaRows.length;
  return (
    <PresSection title="Cash, CD's, Business & Other" badge={formatDollars(total)}>
      <div className="pres-split-col">
        {hasLeft && (
          <div>
            {cashRows.length > 0 && <>
              <div className="pres-col-title">Cash &amp; CD's · {formatDollars(cashCd)}</div>
              <PresTable
                columns={[
                  { label: 'Description',  key: 'description' },
                  { label: 'Market Value', key: '_mv', right: true },
                  { label: 'Int. Rate',    key: '_rate' },
                  { label: 'Owner',        key: 'ownership' },
                ]}
                rows={cashRows.map(r => ({
                  ...r,
                  _mv: formatDollars(num(r.marketValue)),
                  _rate: r.interestRate ? r.interestRate + '%' : '—',
                }))}
                footTotals={{ _mv: formatDollars(cashCd) }}
              />
            </>}
            {plan529Rows.length > 0 && <>
              <div className="pres-col-title pres-col-title-sub">529 Plans · {formatDollars(plan529)}</div>
              <PresTable
                columns={[
                  { label: 'Description',  key: 'description' },
                  { label: 'Market Value', key: '_mv', right: true },
                  { label: 'Owner',        key: 'ownership' },
                ]}
                rows={plan529Rows.map(r => ({ ...r, _mv: formatDollars(num(r.marketValue)) }))}
                footTotals={{ _mv: formatDollars(plan529) }}
              />
            </>}
            {hsaRows.length > 0 && <>
              <div className="pres-col-title pres-col-title-sub">HSA · {formatDollars(hsaVal)}</div>
              <PresTable
                columns={[
                  { label: 'Description',  key: 'description' },
                  { label: 'Market Value', key: '_mv', right: true },
                  { label: 'Owner',        key: 'ownership' },
                ]}
                rows={hsaRows.map(r => ({ ...r, _mv: formatDollars(num(r.marketValue)) }))}
                footTotals={{ _mv: formatDollars(hsaVal) }}
              />
            </>}
          </div>
        )}
        {busRows.length > 0 && (
          <div>
            <div className="pres-col-title">Business &amp; Other · {formatDollars(busOther)}</div>
            <PresTable
              columns={[
                { label: 'Description',  key: 'description' },
                { label: 'Market Value', key: '_mv', right: true },
                { label: 'Cost Basis',   key: '_cb', right: true },
                { label: 'Owner',        key: 'ownership' },
              ]}
              rows={busRows.map(r => ({
                ...r,
                _mv: formatDollars(num(r.marketValue)),
                _cb: num(r.costBasis) ? formatDollars(num(r.costBasis)) : '—',
              }))}
              footTotals={{
                _mv: formatDollars(busOther),
                _cb: sumA(busRows, 'costBasis') ? formatDollars(sumA(busRows, 'costBasis')) : '—',
              }}
            />
          </div>
        )}
      </div>
    </PresSection>
  );
}

function LiabilitiesSection({ rows, totalLiab, totalAssets, netWorth }) {
  const liabPmtTot = sumA(rows, 'payment');
  return (
    <PresSection title="Other Liabilities" badge={formatDollars(totalLiab)}>
      {rows.length > 0 ? (
        <PresTable
          columns={[
            { label: 'Description',  key: 'description' },
            { label: 'Payment (mo)', key: '_pmt',  right: true },
            { label: 'Amount',       key: '_amt',  right: true },
            { label: 'Int. Rate',    key: '_rate', right: true },
            { label: 'Term',         key: 'term' },
          ]}
          rows={rows.map(r => ({
            ...r,
            _pmt:  num(r.payment) ? formatDollars(num(r.payment)) : '—',
            _amt:  formatDollars(num(r.amount)),
            _rate: r.interestRate ? r.interestRate + '%' : '—',
          }))}
          footTotals={{
            _pmt: liabPmtTot ? formatDollars(liabPmtTot) : '—',
            _amt: formatDollars(totalLiab),
          }}
        />
      ) : (
        <p className="pres-empty-note">No liabilities reported.</p>
      )}

      <div className="pres-networth-footer">
        <div className="pres-networth-col">
          {presRow('Total Assets', formatDollars(totalAssets))}
          {presRow('Total Liabilities', formatDollars(totalLiab))}
        </div>
        <div className="pres-networth-total">
          <span className="pres-networth-label">Total Net Worth</span>
          <span className={'pres-networth-value ' + (netWorth >= 0 ? 'positive' : 'negative')}>
            {formatDollars(netWorth)}
          </span>
        </div>
      </div>
    </PresSection>
  );
}

function InsuranceSection({ rows, policies }) {
  const totalPrem    = sumA(policies, 'annualPremium');
  const totalCV      = sumA(policies, 'cashValue');
  const totalBenefit = sumA(policies, 'benefit');
  return (
    <PresSection title="Insurance Information">
      <PresTable
        columns={[
          { label: 'Company',             key: 'company' },
          { label: 'Type',                key: 'type' },
          { label: 'Death/Daily Benefit', key: '_benefit', right: true },
          { label: 'Insured',             key: 'insured' },
          { label: 'Owner',               key: 'owner' },
          { label: 'Policy Date',         key: 'policyDate' },
          { label: 'Annual Premium',      key: '_prem', right: true },
          { label: 'Cash Value',          key: '_cv',   right: true },
          { label: 'Beneficiary',         key: 'beneficiary' },
        ]}
        rows={rows.map(r => ({
          ...r,
          _benefit: num(r.benefit)       ? formatDollars(num(r.benefit))       : '—',
          _cv:      num(r.cashValue)     ? formatDollars(num(r.cashValue))     : '—',
          _prem:    num(r.annualPremium) ? formatDollars(num(r.annualPremium)) : '—',
        }))}
        footTotals={{
          _benefit: totalBenefit ? formatDollars(totalBenefit) : '—',
          _prem:    formatDollars(totalPrem),
          _cv:      formatDollars(totalCV),
        }}
      />
    </PresSection>
  );
}

function IncomeTaxSection({ empRows, ssRows, pensRows, otherRows, empInc, ssInc, pensInc, otherInc, totalIncome, taxes, fedTax, stTax, ficaTax, totalTax }) {
  const hasIncome = empRows.length || ssRows.length || pensRows.length || otherRows.length;
  const hasTaxSummary = num(taxes.capLossCarryForward) || num(taxes.taxableIncome) || num(taxes.standardItemDeduction) || totalTax;

  return (
    <PresSection title="Income & Tax Information" badge={formatDollars(totalIncome)}>
      <div className="pres-income-layout">
        <div className="pres-income-main">
          {empRows.length > 0 && <>
            <div className="pres-col-title">Employment Income · {formatDollars(empInc)}</div>
            <PresTable
              columns={[
                { label: 'Owner',           key: 'owner' },
                { label: 'Description',     key: 'description' },
                { label: 'Annual Amount',   key: '_amt',  right: true },
                { label: 'Retirement Date', key: 'retirementDate' },
                { label: 'COLA %',          key: '_cola', right: true },
              ]}
              rows={empRows.map(r => ({ ...r, _amt: formatDollars(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' }))}
              footTotals={{ _amt: formatDollars(empInc) }}
            />
          </>}

          {ssRows.length > 0 && <>
            <div className="pres-col-title pres-col-title-sub">Social Security · {formatDollars(ssInc)}</div>
            <PresTable
              columns={[
                { label: 'Owner',         key: 'owner' },
                { label: 'Starting Age',  key: 'startingAge', right: true },
                { label: 'Annual Amount', key: '_amt',        right: true },
                { label: 'Full Ret. Age', key: 'fullRetAge',  right: true },
                { label: 'COLA %',        key: '_cola',       right: true },
              ]}
              rows={ssRows.map(r => ({ ...r, _amt: formatDollars(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' }))}
              footTotals={{ _amt: formatDollars(ssInc) }}
            />
          </>}

          {pensRows.length > 0 && <>
            <div className="pres-col-title pres-col-title-sub">Pension · {formatDollars(pensInc)}</div>
            <PresTable
              columns={[
                { label: 'Owner',            key: 'owner' },
                { label: 'Start Date',       key: 'startDate' },
                { label: 'Annual Amount',    key: '_amt',  right: true },
                { label: 'Survivor Benefit', key: '_surv', right: true },
                { label: 'COLA %',           key: '_cola', right: true },
              ]}
              rows={pensRows.map(r => ({ ...r, _amt: formatDollars(num(r.annualAmount)), _surv: r.survivorBenefit ? r.survivorBenefit + '%' : '—', _cola: r.cola ? r.cola + '%' : '—' }))}
              footTotals={{ _amt: formatDollars(pensInc) }}
            />
          </>}

          {otherRows.length > 0 && <>
            <div className="pres-col-title pres-col-title-sub">Other Income · {formatDollars(otherInc)}</div>
            <PresTable
              columns={[
                { label: 'Owner',         key: 'owner' },
                { label: 'Description',   key: 'description' },
                { label: 'Annual Amount', key: '_amt', right: true },
                { label: 'Start',         key: 'startDate' },
                { label: 'End',           key: 'endDate' },
                { label: 'COLA %',        key: '_cola', right: true },
              ]}
              rows={otherRows.map(r => ({ ...r, _amt: formatDollars(num(r.annualAmount)), _cola: r.cola ? r.cola + '%' : '—' }))}
              footTotals={{ _amt: formatDollars(otherInc) }}
            />
          </>}

          {hasIncome > 0 && (
            <div className="pres-income-total">
              <span>Current Total</span>
              <span>{formatDollars(totalIncome)}</span>
            </div>
          )}
        </div>

        {hasTaxSummary > 0 && (
          <aside className="pres-tax-panel">
            <div className="pres-tax-panel-title">Prior Year Tax Summary</div>
            {num(taxes.capLossCarryForward) > 0   && <TaxRow label="Cap Loss Carry Forward" value={formatDollars(num(taxes.capLossCarryForward))} />}
            {num(taxes.taxableIncome) > 0         && <TaxRow label="Taxable Income"         value={formatDollars(num(taxes.taxableIncome))} />}
            {num(taxes.standardItemDeduction) > 0 && <TaxRow label="Stnd/Item Deduction"    value={formatDollars(num(taxes.standardItemDeduction))} />}
            {totalTax > 0 && <>
              <div className="pres-tax-subhead">Taxes Paid</div>
              {fedTax  > 0 && <TaxRow label="Federal Tax" value={formatDollars(fedTax)} />}
              {stTax   > 0 && <TaxRow label="State Tax"   value={formatDollars(stTax)} />}
              {ficaTax > 0 && <TaxRow label="FICA Tax"    value={formatDollars(ficaTax)} />}
              <TaxRow label="Total Taxes Paid" value={formatDollars(totalTax)} emph />
              {totalIncome > 0 && <TaxRow label="Effective Rate" value={((totalTax / totalIncome) * 100).toFixed(1) + '%'} />}
            </>}
          </aside>
        )}
      </div>
    </PresSection>
  );
}

function TaxRow({ label, value, emph }) {
  return (
    <div className={'pres-tax-row' + (emph ? ' is-total' : '')}>
      <span>{label}</span>
      <span>{value}</span>
    </div>
  );
}

function ExpensesCashFlowSection({ livingExpField, expRows, totalIncome, totalTax, ficaTax, totalSavings, debtPayments, livingExp, netCF, savingsRatio, dtiRatio }) {
  const fedStateTax = totalTax - ficaTax;
  const cfLines = [
    { label: 'Gross Income',                  val: totalIncome,  minus: false },
    { label: 'Less: Federal & State Tax',     val: fedStateTax,  minus: true, skip: fedStateTax === 0 },
    { label: 'Less: FICA Tax',                val: ficaTax,      minus: true, skip: ficaTax === 0 },
    { label: 'Less: Savings & Contributions', val: totalSavings, minus: true, skip: totalSavings === 0 },
    { label: 'Less: Debt Payments',           val: debtPayments, minus: true, skip: debtPayments === 0 },
    { label: 'Less: Living Expenses',         val: livingExp,    minus: true, skip: livingExp === 0 },
  ].filter(l => !l.skip);

  return (
    <PresSection title="Expenses & Cash Flow">
      {livingExpField > 0 && <>
        <div className="pres-subsec-label">Living Expenses</div>
        <InfoGrid items={[{ label: 'Annual Amount', value: formatDollars(livingExpField) }]} />
      </>}

      {expRows.length > 0 && <>
        <div className="pres-subsec-label">Other Expenses &amp; Personal Liabilities</div>
        <PresTable
          columns={[
            { label: 'Description',   key: 'description' },
            { label: 'Annual Amount', key: '_amt',  right: true },
            { label: 'Start',         key: 'startDate' },
            { label: 'End',           key: 'endDate' },
            { label: 'COLA %',        key: '_cola', right: true },
            { label: 'Notes',         key: 'notes' },
          ]}
          rows={expRows.map(r => ({ ...r, _amt: formatDollars(num(r.amount)), _cola: r.cola ? r.cola + '%' : '—' }))}
          footTotals={{ _amt: formatDollars(expRows.reduce((s, r) => s + num(r.amount), 0)) }}
        />
      </>}

      <div className="pres-subsec-label">Annual Cash Flow</div>
      <div className="pres-cf-layout">
        <table className="pres-cf-table">
          <tbody>
            {cfLines.map((line, i) => (
              <tr key={i}>
                <td className="cf-label">{line.label}</td>
                <td className={'cf-value' + (line.minus && line.val > 0 ? ' cf-negative' : '')}>
                  {(line.minus && line.val > 0 ? '(' : '') + formatDollars(line.val) + (line.minus && line.val > 0 ? ')' : '')}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className={'pres-cf-net' + (netCF >= 0 ? ' is-positive' : ' is-negative')}>
          <div className="pres-cf-net-label">Net Cash Flow</div>
          <div className="pres-cf-net-value">{formatDollars(netCF)}</div>
          <div className="pres-cf-net-hint">{netCF >= 0 ? 'Annual surplus' : 'Annual shortfall'}</div>

          {totalIncome > 0 && (
            <div className="pres-cf-ratios">
              <Ratio label="Savings Ratio"        pct={savingsRatio} hint="Total Savings / Gross Income" />
              <Ratio label="Debt-to-Income Ratio" pct={dtiRatio}     hint="Annual Debt Payments / Gross Income" />
            </div>
          )}
        </div>
      </div>
    </PresSection>
  );
}

function Ratio({ label, pct, hint }) {
  return (
    <div className="pres-ratio">
      <div className="pres-ratio-label">{label}</div>
      <div className="pres-ratio-value">{pct.toFixed(1)}%</div>
      {hint && <div className="pres-ratio-hint">{hint}</div>}
    </div>
  );
}
