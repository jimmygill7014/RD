export default function PQScreen() {
  return (
    <section className="screen is-active">
      <section className="page-title">
        <h1>Personal Financial Questionnaire</h1>
      </section>
      <form id="pq-form" autoComplete="off">
        <div className="pq-layout">
          <aside className="pq-notes-panel"></aside>
          <div className="form-col">
            {/* TODO: port form sections from app.js into <FormSection /> components */}
            <p className="section-note">Form sections will render here.</p>
          </div>
          <aside className="pq-dashboard"></aside>
        </div>
        <footer className="actions">
          <button type="button" className="btn btn-light">Clear Form</button>
          <button type="submit" className="btn btn-primary">Save</button>
        </footer>
      </form>
    </section>
  );
}
