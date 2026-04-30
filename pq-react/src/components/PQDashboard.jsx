import { useStore } from '../store/StoreContext.jsx';
import { getDashboardMetrics, formatDollars } from '../store/selectors.js';
import TaxTriangleSVG from './TaxTriangleSVG.jsx';

function pct(n) {
  if (!isFinite(n)) return '0%';
  return n.toFixed(1) + '%';
}

function dtiColorClass(dti) {
  if (dti > 40) return 'negative';
  if (dti > 20) return 'warning';
  return 'positive';
}

function savingsRatioColorClass(ratio) {
  if (ratio >= 20) return 'positive';
  if (ratio >= 10) return 'warning';
  return 'negative';
}

function AssetRow({ label, value, swatch }) {
  return (
    <div className="bs-asset-row">
      <div className="bs-asset-left">
        <span className={`bs-swatch bs-swatch-${swatch}`} />
        <span className="bs-asset-label">{label}</span>
      </div>
      <strong className="bs-asset-value">{formatDollars(value)}</strong>
    </div>
  );
}

export default function PQDashboard() {
  const { data } = useStore();
  const m = getDashboardMetrics(data);

  return (
    <>
      {/* Combined Balance Sheet + Tax Triangle card */}
      <div className="dashboard-card balance-triangle-card">
        <div className="balance-triangle-inner">
          <div className="balance-col">
            <h3>Balance Sheet</h3>
            <div className="bs-breakdown">
              <AssetRow label="Investable Assets" value={m.investableAssets} swatch="invest" />
              <AssetRow label="Real Estate"        value={m.realEstateAssets} swatch="realestate" />
              <AssetRow label="Business / Other"   value={m.businessOtherAssets} swatch="business" />
            </div>
            <div className="bs-total-assets">
              <span className="bs-total-label">Total Assets</span>
              <strong className="bs-total-value">{formatDollars(m.totalAssets)}</strong>
            </div>
            <div className="bs-liab-line">
              <span>Total Liabilities</span>
              <strong className="metric-value">{formatDollars(m.totalLiabilities)}</strong>
            </div>
            <div className="bs-net-worth">
              <span>Net Worth</span>
              <strong className={`metric-value ${m.totalNetWorth >= 0 ? 'positive' : 'negative'}`}>
                {formatDollars(m.totalNetWorth)}
              </strong>
            </div>
          </div>
          <div className="triangle-col">
            <div className="tax-triangle-wrapper">
              <TaxTriangleSVG
                taxFree={m.taxFree}
                taxable={m.taxable}
                taxDeferred={m.taxDeferred}
              />
            </div>
          </div>
        </div>
      </div>

      {/* Yearly Cash Flow card */}
      <div className="dashboard-card">
        <h3>Yearly Cash Flow</h3>
        <ul className="metric-list">
          <li>
            <span>Income</span>
            <strong className="metric-value">{formatDollars(m.totalIncome)}</strong>
          </li>
          <li>
            <span>Savings</span>
            <strong className="metric-value">{formatDollars(m.totalSavings)}</strong>
          </li>
          <li>
            <span>Liability Payments</span>
            <strong className="metric-value">{formatDollars(m.liabilityPayments)}</strong>
          </li>
          <li>
            <span>Taxes</span>
            <strong className="metric-value">{formatDollars(m.totalTax)}</strong>
          </li>
          <li>
            <span>Non-Liability Expenses</span>
            <strong className="metric-value">{formatDollars(m.nonLiabilityExpenses)}</strong>
          </li>
          <li style={{ borderTop: '1px solid var(--border)', paddingTop: 6, marginTop: 4 }}>
            <span style={{ fontWeight: 700 }}>Cash Flow</span>
            <strong className={`metric-value ${m.cashFlow >= 0 ? 'positive' : 'negative'}`}>
              {formatDollars(m.cashFlow)}
            </strong>
          </li>
          <li style={{ borderTop: '1px solid var(--border)', paddingTop: 6, marginTop: 4 }}>
            <span>Debt-to-Income</span>
            <strong className={`metric-value ${dtiColorClass(m.dti)}`}>{pct(m.dti)}</strong>
          </li>
          <li>
            <span>Savings Ratio</span>
            <strong className={`metric-value ${savingsRatioColorClass(m.savingsRatio)}`}>
              {pct(m.savingsRatio)}
            </strong>
          </li>
        </ul>
      </div>
    </>
  );
}
