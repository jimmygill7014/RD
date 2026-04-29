import { useStore } from '../store/StoreContext.jsx';
import Field from './Field.jsx';

export default function ConditionalBlock({ block, sectionId, onToggleOff, children, headerSlot }) {
  const { data, update } = useStore();
  const flag = !!data?._flags?.[block.flag];

  const toggle = () => {
    const next = !flag;
    update(`_flags.${block.flag}`, next);
    if (!next && onToggleOff) onToggleOff();
  };

  const button = (
    <button type="button" className="btn-mini" onClick={toggle}>
      {flag ? block.hideLabel : block.toggleLabel}
    </button>
  );

  if (headerSlot) {
    return headerSlot({ flag, button, content: flag ? children : null });
  }

  return (
    <div className="subsection">
      {button}
      <div
        className={'conditional-section' + (flag ? '' : ' hidden')}
        style={{ marginTop: 10 }}
      >
        {flag && children}
      </div>
    </div>
  );
}
