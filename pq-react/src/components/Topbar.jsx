import { useStore } from '../store/StoreContext.jsx';

const STATUS_LABEL = {
  idle: '',
  saving: 'Saving…',
  saved: 'Saved ✓',
};

const STATUS_CLASS = {
  idle: 'autosave-status',
  saving: 'autosave-status is-saving',
  saved: 'autosave-status is-saved',
};

export default function Topbar({ onPresentation, onPlan, onOpenConsole }) {
  const { autosaveStatus, saveNow, reset } = useStore();

  const handleReset = () => {
    if (window.confirm('Reset all data and start over?')) reset();
  };

  return (
    <header className="topbar">
      <a className="brand" href="#" aria-label="Pure Financial Advisors">
        <span className="brand-text">Pure Financial Advisors</span>
      </a>
      <div className="topbar-actions">
        <span className={STATUS_CLASS[autosaveStatus]} aria-live="polite">
          {STATUS_LABEL[autosaveStatus]}
        </span>
        <button type="button" className="btn btn-primary" onClick={saveNow}>
          Save
        </button>
        <button type="button" className="btn btn-ghost" onClick={onPresentation}>
          Presentation Mode
        </button>
        <button type="button" className="btn btn-ghost" onClick={onPlan}>
          Plan Mode
        </button>
        <button type="button" className="btn btn-ghost" onClick={onOpenConsole}>
          Data Console
        </button>
        <button
          type="button"
          className="btn btn-ghost btn-ghost-danger"
          onClick={handleReset}
        >
          Reset Data
        </button>
      </div>
    </header>
  );
}
