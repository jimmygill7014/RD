import { formatDollars } from '../store/selectors.js';

const TOP = { cx: 160, cy: 80 };
const BL  = { cx: 82,  cy: 210 };
const BR  = { cx: 238, cy: 210 };
const R   = 62;

// Dashed connector lines stay inside the circle radii.
const CONNECTORS = [
  [TOP.cx, TOP.cy + R, BL.cx + R * 0.65, BL.cy - R * 0.65],
  [BL.cx + R, BL.cy, BR.cx - R, BR.cy],
  [BR.cx - R * 0.65, BR.cy - R * 0.65, TOP.cx, TOP.cy + R],
];

function CircleNode({ cx, cy, gradId, label, value, percent, filterId }) {
  return (
    <g>
      <circle cx={cx} cy={cy} r={R} filter={`url(#${filterId})`} fill="white" />
      <circle cx={cx} cy={cy} r={R} fill={`url(#${gradId})`} />
      <circle cx={cx} cy={cy} r={R - 3} fill="none" stroke="rgba(255,255,255,0.2)" strokeWidth="1" />
      <text x={cx} y={cy - 16} textAnchor="middle" className="tri-node-label">{label}</text>
      <text x={cx} y={cy + 4}  textAnchor="middle" className="tri-node-value">{formatDollars(value)}</text>
      <text x={cx} y={cy + 24} textAnchor="middle" className="tri-node-pct">{percent}%</text>
    </g>
  );
}

export default function TaxTriangleSVG({ taxFree, taxable, taxDeferred, idSuffix = '' }) {
  const total = taxFree + taxable + taxDeferred;
  const pct = v => total > 0 ? Math.round((v / total) * 100) : 0;
  const filterId = 'triShadow' + idSuffix;
  const gradFreeId = 'gradFree' + idSuffix;
  const gradTaxableId = 'gradTaxable' + idSuffix;
  const gradDeferredId = 'gradDeferred' + idSuffix;

  return (
    <svg viewBox="0 0 320 280" preserveAspectRatio="xMidYMid meet">
      <defs>
        <filter id={filterId} x="-20%" y="-20%" width="140%" height="140%">
          <feGaussianBlur in="SourceAlpha" stdDeviation="3" />
          <feOffset dx="0" dy="2" />
          <feMerge>
            <feMergeNode />
            <feMergeNode in="SourceGraphic" />
          </feMerge>
        </filter>
        <linearGradient id={gradFreeId} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%"   stopColor="#1a8fc4" />
          <stop offset="100%" stopColor="#1578a8" />
        </linearGradient>
        <linearGradient id={gradTaxableId} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%"   stopColor="#2b5ea7" />
          <stop offset="100%" stopColor="#1e4a8a" />
        </linearGradient>
        <linearGradient id={gradDeferredId} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%"   stopColor="#0d6efd" />
          <stop offset="100%" stopColor="#0a58ca" />
        </linearGradient>
      </defs>

      {CONNECTORS.map(([x1, y1, x2, y2], i) => (
        <line key={i} x1={x1} y1={y1} x2={x2} y2={y2}
              stroke="#c0cdd8" strokeWidth="1.5" strokeDasharray="4,3" />
      ))}

      <CircleNode {...TOP} gradId={gradFreeId}     filterId={filterId} label="TAX-FREE"     value={taxFree}     percent={pct(taxFree)} />
      <CircleNode {...BL}  gradId={gradTaxableId}  filterId={filterId} label="TAXABLE"      value={taxable}     percent={pct(taxable)} />
      <CircleNode {...BR}  gradId={gradDeferredId} filterId={filterId} label="TAX-DEFERRED" value={taxDeferred} percent={pct(taxDeferred)} />
    </svg>
  );
}
