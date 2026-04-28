export default function Topbar({ onPresentation, onPlan, onOpenConsole }) {
  return (
    <header className="topbar">
      <a className="brand" href="#" aria-label="Pure Financial Advisors">
        <span className="brand-text">Pure Financial Advisors</span>
      </a>
      <div className="topbar-actions">
        <span className="autosave-status" aria-live="polite"></span>
        <button type="button" className="btn btn-primary">Save</button>
        <button type="button" className="btn btn-ghost" onClick={onPresentation}>
          Presentation Mode
        </button>
        <button type="button" className="btn btn-ghost" onClick={onPlan}>
          Plan Mode
        </button>
        <button type="button" className="btn btn-ghost" onClick={onOpenConsole}>
          Data Console
        </button>
        <button type="button" className="btn btn-ghost btn-ghost-danger">
          Reset Data
        </button>
      </div>
    </header>
  );
}
