import { pqSections } from '../schema/pqSections.js';
import FormSection from '../components/FormSection.jsx';

export default function PQScreen() {
  return (
    <section className="screen is-active">
      <section className="page-title">
        <h1>Personal Financial Questionnaire</h1>
      </section>
      <form id="pq-form" autoComplete="off" onSubmit={e => e.preventDefault()}>
        <div className="pq-layout">
          <aside className="pq-notes-panel"></aside>
          <div className="form-col">
            {pqSections.map(s => (
              <FormSection key={s.id} section={s} />
            ))}
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
