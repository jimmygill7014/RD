import { useStore } from '../store/StoreContext.jsx';

export default function NotesDrawer({ open, onToggle }) {
  const { data, update } = useStore();
  const goals = data._goals ?? '';
  const notes = data._advisorNotes ?? '';

  return (
    <aside className={'notes-drawer' + (open ? ' is-open' : '')} aria-hidden={!open}>
      <button
        type="button"
        className="notes-drawer-tab"
        aria-label="Open notes and goals"
        onClick={onToggle}
      >
        <span className="notes-drawer-tab-label">Goals &amp; Notes</span>
        <span className="notes-drawer-tab-arrow" aria-hidden="true">›</span>
      </button>
      <div className="notes-drawer-panel">
        <header className="notes-drawer-header">
          <h2>Goals &amp; Advisor Notes</h2>
          <button type="button" className="btn btn-light" onClick={onToggle} aria-label="Close">
            Close
          </button>
        </header>
        <div className="notes-drawer-body">
          <label className="notes-drawer-label" htmlFor="drawer-goals">Client Goals</label>
          <textarea
            id="drawer-goals"
            className="notes-drawer-textarea"
            placeholder="Retirement, college, travel, legacy..."
            value={goals}
            onChange={e => update('_goals', e.target.value)}
          />
          <label className="notes-drawer-label" htmlFor="drawer-notes">Advisor Notes</label>
          <textarea
            id="drawer-notes"
            className="notes-drawer-textarea"
            placeholder="Free-form notes..."
            value={notes}
            onChange={e => update('_advisorNotes', e.target.value)}
          />
        </div>
      </div>
    </aside>
  );
}
