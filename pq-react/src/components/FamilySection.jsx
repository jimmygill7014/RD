import { useStore } from '../store/StoreContext.jsx';
import Field from './Field.jsx';
import DataTable from './DataTable.jsx';
import ConditionalBlock from './ConditionalBlock.jsx';

export default function FamilySection({ section }) {
  const { data, update, stripKeysMatching } = useStore();
  const sectionData = data[section.id] || {};
  const flags = data._flags || {};

  const renderField = f => (
    <Field
      key={f.key}
      field={f}
      value={sectionData[f.key]}
      onChange={val => update(`${section.id}.${f.key}`, val)}
    />
  );

  const spouseBlock = section.conditionalBlocks?.find(b => b.id === 'spouse-block');
  const childrenBlock = section.afterConditional?.find(b => b.id === 'children-block');

  const handleSpouseOff = () => {
    // Mirror legacy clearSpouseData: strip any client2* keys + family.yearsMarried
    stripKeysMatching(k => typeof k === 'string' && k.startsWith('client2'));
    update('family.yearsMarried', undefined);
  };

  return (
    <section className={`form-section theme-${section.colorTheme || 'default'}`}>
      <div className="section-header">
        <h2>{section.title}</h2>
      </div>
      <div className="section-content">
        <div className="family-split">
          {/* Client column */}
          <div className="family-split-col family-split-client-col">
            <div className="family-split-label">Client</div>
            <div className="grid">{(section.fields || []).map(renderField)}</div>
          </div>

          {/* Spouse column */}
          <div className="family-split-col family-split-spouse-col">
            {spouseBlock && (
              <ConditionalBlock
                block={spouseBlock}
                sectionId={section.id}
                onToggleOff={handleSpouseOff}
                headerSlot={({ flag, button, content }) => (
                  <>
                    <div
                      className={
                        'family-split-label family-split-label--with-action' +
                        (flag ? '' : ' family-split-label--collapsed')
                      }
                    >
                      {flag && 'Spouse / Partner'}
                      <span
                        className={
                          'family-split-header-btn ' +
                          (flag ? 'family-split-header-btn--remove' : 'family-split-header-btn--add')
                        }
                      >
                        {button}
                      </span>
                    </div>
                    {flag && (
                      <div className="subsection subsection-spouse">
                        <div className="grid">{spouseBlock.fields.map(renderField)}</div>
                      </div>
                    )}
                  </>
                )}
              >
                {/* content rendered via headerSlot */}
              </ConditionalBlock>
            )}
          </div>
        </div>

        {/* Children / grandchildren block */}
        {childrenBlock && (
          <ConditionalBlock block={childrenBlock} sectionId={section.id}>
            {(childrenBlock.tables || []).map(t => (
              <DataTable key={t.key} sectionId={section.id} tableDef={t} />
            ))}
            {childrenBlock.summaryFields?.length > 0 && (
              <div className="grid" style={{ marginTop: 12 }}>
                {childrenBlock.summaryFields.map(renderField)}
              </div>
            )}
          </ConditionalBlock>
        )}

        {/* Footer fields (additional family info) */}
        {section.footerFields?.length > 0 && (
          <div className="grid" style={{ marginTop: 12 }}>
            {section.footerFields.map(renderField)}
          </div>
        )}
      </div>
    </section>
  );
}
