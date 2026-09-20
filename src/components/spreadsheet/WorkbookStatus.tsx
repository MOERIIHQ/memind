import type { Workbook } from '../../types/spreadsheet';
import Icon from '../Icon';
import './WorkbookStatus.css';

interface WorkbookStatusProps {
  workbook: Workbook | null;
  error: string | null;
  onClearError: () => void;
}

const WorkbookStatus = ({ workbook, error }: WorkbookStatusProps) => {
  if (error) {
    return (
      <div className="workbook-status workbook-status--error" role="alert">
        <div className="workbook-status__icon">
          <Icon name="alertCircle" size={20} />
        </div>
        <div className="workbook-status__content">
          <span className="workbook-status__title">Something went wrong</span>
          <p className="workbook-status__message">{error}</p>
        </div>
      </div>
    );
  }

  if (!workbook) {
    return null;
  }

  const activeSheet = workbook.sheets[workbook.activeSheetIndex];

  return (
    <div className="workbook-status">
      <div className="workbook-status__item">
        <Icon name="file-spreadsheet" size={16} />
        <span className="workbook-status__label">File</span>
        <span className="workbook-status__value">{workbook.metadata.fileName}</span>
      </div>
      <div className="workbook-status__item">
        <Icon name="table" size={16} />
        <span className="workbook-status__label">Sheet</span>
        <span className="workbook-status__value">
          {activeSheet ? activeSheet.name : '—'}
        </span>
      </div>
      <div className="workbook-status__item">
        <Icon name="grid" size={16} />
        <span className="workbook-status__label">Dimensions</span>
        <span className="workbook-status__value">
          {activeSheet ? `${activeSheet.colCount} × ${activeSheet.rowCount}` : '—'}
        </span>
      </div>
      <div className="workbook-status__item">
        <Icon name="file" size={16} />
        <span className="workbook-status__label">Size</span>
        <span className="workbook-status__value">{formatFileSize(workbook.metadata.fileSize)}</span>
      </div>
    </div>
  );
};

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default WorkbookStatus;
