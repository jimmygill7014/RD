export default function PlanScreen({ onBack }) {
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
          <button type="button" className="btn btn-primary">Send to Planning</button>
        </div>
      </div>
      <div className="plan-layout">
        <aside className="plan-summary"></aside>
        <div className="plan-form-wrap">
          <form autoComplete="off">
            <div>{/* TODO: port plan form sections */}</div>
          </form>
        </div>
      </div>
    </section>
  );
}
