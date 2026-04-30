import { useState } from 'react';
import { useStore } from '../store/StoreContext.jsx';
import { getDashboardMetrics, formatDollars } from '../store/selectors.js';
import { planSections } from '../schema/planSections.js';
import PlanField from '../components/PlanField.jsx';

function getSmartPrefills(data, m) {
  let cashFlow = '';
  if (m.totalIncome > 0) {
    if (m.cashFlow > 50) cashFlow = 'Surplus';
    else if (m.cashFlow < -50) cashFlow = 'Deficit';
    else cashFlow = 'Break-Even';
  }
  return {
    clientRetirementDate: data?.employment?.client1?.retirementDate || '',
    spouseRetirementDate: data?.employment?.client2?.retirementDate || '',
    currentCashFlow: cashFlow,
  };
}

function PlanSummaryPanel({ data, m }) {
  const family = data.family || {};
  const c1 = [family.client1FirstName, family.client1LastName].filter(Boolean).join(' ');
  const c2 = [family.client2FirstName, family.client2LastName].filter(Boolean).join(' ');

  return (
    <div className="plan-summary-inner">
      <h3>Client Snapshot</h3>
      <div className="plan-sum-name">{c1 || 'Client'}{c2 ? ` & ${c2}` : ''}</div>
      <ul className="metric-list">
        <li><span>Net Worth</span>
          <strong className={`metric-value ${m.totalNetWorth >= 0 ? 'positive' : 'negative'}`}>
            {formatDollars(m.totalNetWorth)}
          </strong>
        </li>
        <li><span>Total Assets</span><strong className="metric-value">{formatDollars(m.totalAssets)}</strong></li>
        <li><span>Total Liabilities</span><strong className="metric-value">{formatDollars(m.totalLiabilities)}</strong></li>
        <li><span>Annual Income</span><strong className="metric-value">{formatDollars(m.totalIncome)}</strong></li>
        <li><span>Annual Taxes</span><strong className="metric-value">{formatDollars(m.totalTax)}</strong></li>
        <li><span>Annual Savings</span><strong className="metric-value">{formatDollars(m.totalSavings)}</strong></li>
        <li style={{ borderTop: '1px solid var(--border)', paddingTop: 6, marginTop: 4 }}>
          <span style={{ fontWeight: 700 }}>Cash Flow</span>
          <strong className={`metric-value ${m.cashFlow >= 0 ? 'positive' : 'negative'}`}>
            {formatDollars(m.cashFlow)}
          </strong>
        </li>
      </ul>
      {data._goals && (
        <>
          <h3 style={{ marginTop: 16 }}>Goals</h3>
          <div className="plan-sum-notes">{data._goals}</div>
        </>
      )}
    </div>
  );
}

export default function PlanScreen({ onBack }) {
  const { data, update } = useStore();
  const m = getDashboardMetrics(data);
  const [sentFlash, setSentFlash] = useState(false);

  const planOrder = data.planOrder || {};
  const prefills = getSmartPrefills(data, m);

  const valueFor = field => {
    const stored = planOrder[field.key];
    if (stored !== undefined) return stored;
    return prefills[field.key] ?? '';
  };

  const setField = (key, val) => update(`planOrder.${key}`, val);

  const handleSend = () => {
    update('_workflow.planOrderSavedAt', new Date().toISOString());
    setSentFlash(true);
    setTimeout(() => setSentFlash(false), 1500);
  };

  return (
    <section className="screen is-active">
      <div className="plan-nav no-print">
        <div className="plan-nav-brand brand" aria-label="Pure Financial Advisors — Plan Order">
          <span className="brand-text">Pure Financial Advisors</span>
          <span className="brand-divider" aria-hidden="true"></span>
          <span className="brand-context">Plan Order</span>
        </div>
        <div className="plan-nav-actions">
          <button type="button" className="btn btn-light" onClick={onBack}>
            ← Back to PQ
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSend}
            disabled={sentFlash}
          >
            {sentFlash ? 'Sent ✓' : 'Send to Planning'}
          </button>
        </div>
      </div>

      <div className="plan-layout">
        <aside className="plan-summary">
          <PlanSummaryPanel data={data} m={m} />
        </aside>
        <div className="plan-form-wrap">
          <form autoComplete="off" onSubmit={e => e.preventDefault()}>
            <div>
              {planSections.map(section => (
                <div key={section.id} className={`plan-form-card plan-card-${section.color}`}>
                  <div className="plan-card-header">{section.title}</div>
                  <div className="plan-card-body">
                    {section.fields.map(field => (
                      <PlanField
                        key={field.key}
                        field={field}
                        value={valueFor(field)}
                        onChange={val => setField(field.key, val)}
                      />
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </form>
        </div>
      </div>
    </section>
  );
}
