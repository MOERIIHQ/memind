import type { CellAddress } from '../../types/spreadsheet';
import Button from '../Button';
import Icon from '../Icon';
import { useRef } from 'react';
import './SpreadsheetToolbar.css';

interface SpreadsheetToolbarProps {
  fileName: string;
  activeSheetName: string;
  selectedCell: CellAddress | null;
  uploadInputId: string;
  onExport: () => void;
  onExportCsv: () => void;
  onFileInputChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
}

const SpreadsheetToolbar = ({
  fileName,
  activeSheetName,
  selectedCell,
  uploadInputId,
  onExport,
  onExportCsv,
  onFileInputChange,
}: SpreadsheetToolbarProps) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const getSelectedCellLabel = (): string => {
    if (!selectedCell) return '-';
    const colLetter = String.fromCharCode(65 + selectedCell.col);
    return `${colLetter}${selectedCell.row + 1}`;
  };

  return (
    <div className="spreadsheet-toolbar">
      <div className="spreadsheet-toolbar__left">
        <label htmlFor={uploadInputId} className="spreadsheet-toolbar__upload-label">
          <input
            ref={inputRef}
            id={uploadInputId}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={onFileInputChange}
            style={{ display: 'none' }}
            tabIndex={-1}
            aria-hidden="true"
          />
          <Button variant="primary" size="md" leftIcon={<Icon name="file" size={16} />}>
            Upload
          </Button>
        </label>
        <div className="spreadsheet-toolbar__file-info">
          <span className="spreadsheet-toolbar__file-name" title={fileName}>
            {fileName || 'Untitled'}
          </span>
          <span className="spreadsheet-toolbar__sheet-badge">
            Sheet: {activeSheetName}
          </span>
        </div>
      </div>

      <div className="spreadsheet-toolbar__center">
        <div className="spreadsheet-toolbar__cell-address">
          <span className="spreadsheet-toolbar__cell-label">Cell</span>
          <code className="spreadsheet-toolbar__cell-value">{getSelectedCellLabel()}</code>
        </div>
      </div>

      <div className="spreadsheet-toolbar__right">
        <Button
          variant="ghost"
          size="sm"
          leftIcon={<Icon name="download" size={14} />}
          onClick={onExportCsv}
          title="Download as CSV"
        >
          CSV
        </Button>
        <Button
          variant="ghost"
          size="sm"
          leftIcon={<Icon name="download" size={14} />}
          onClick={onExport}
          title="Download as Excel"
        >
          XLSX
        </Button>
      </div>
    </div>
  );
};

export default SpreadsheetToolbar;
