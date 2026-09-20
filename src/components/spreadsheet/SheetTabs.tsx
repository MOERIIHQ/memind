import type { Worksheet } from '../../types/spreadsheet';
import Icon from '../Icon';
import './SheetTabs.css';

interface SheetTabsProps {
  sheets: Worksheet[];
  activeSheetIndex: number;
  onSheetSelect: (index: number) => void;
  onSheetAdd?: () => void;
}

const SheetTabs = ({
  sheets,
  activeSheetIndex,
  onSheetSelect,
  onSheetAdd,
}: SheetTabsProps) => {
  return (
    <div className="sheet-tabs" role="tablist" aria-label="Sheets">
      <div className="sheet-tabs__list">
        {sheets.map((sheet, index) => (
          <button
            key={sheet.name}
            type="button"
            className={`sheet-tabs__tab ${index === activeSheetIndex ? 'sheet-tabs__tab--active' : ''}`}
            onClick={() => onSheetSelect(index)}
            role="tab"
            aria-selected={index === activeSheetIndex}
            aria-controls={`sheet-${index}`}
            id={`tab-${index}`}
            title={sheet.name}
          >
            <span className="sheet-tabs__tab-name">{sheet.name}</span>
            {sheet.protected && (
              <span className="sheet-tabs__protected-icon" title="Protected sheet" aria-label="Protected">
                <Icon name="check" size={12} />
              </span>
            )}
          </button>
        ))}
      </div>
      {onSheetAdd && (
        <button
          type="button"
          className="sheet-tabs__add-btn"
          onClick={onSheetAdd}
          aria-label="Add sheet"
          title="Add sheet"
        >
          <Icon name="plus" size={14} />
        </button>
      )}
    </div>
  );
};

export default SheetTabs;
