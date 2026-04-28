export default function DataConsole({ open, onClose, data = {} }) {
  if (!open) return null;
  return (
    <aside className="data-console is-open" aria-hidden={!open}>
      <header className="data-console-header">
        <h2>Data Console</h2>
        <button type="button" className="btn btn-light" onClick={onClose}>
          Close
        </button>
      </header>
      <p className="section-note">Central data store (read-only debug view)</p>
      <pre>{JSON.stringify(data, null, 2)}</pre>
    </aside>
  );
}
