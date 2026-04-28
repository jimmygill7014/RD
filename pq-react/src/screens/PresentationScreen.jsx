export default function PresentationScreen({ onBack }) {
  return (
    <section className="screen is-active">
      <div className="pres-nav no-print">
        <div className="pres-nav-brand brand" aria-label="Pure Financial Advisors">
          <span className="brand-text">Pure Financial Advisors</span>
        </div>
        <div className="pres-nav-actions">
          <button type="button" className="btn btn-light" onClick={onBack}>
            ← Edit
          </button>
          <button type="button" className="btn btn-primary" onClick={() => window.print()}>
            Print / Save PDF
          </button>
        </div>
      </div>
      <div className="pres-content">
        {/* TODO: port presentation rendering from app.js */}
      </div>
    </section>
  );
}
