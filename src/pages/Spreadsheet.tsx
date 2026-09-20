import { useState, useCallback, useRef } from 'react';
import type { CellAddress, CellValue, Workbook, Worksheet } from '../types/spreadsheet';
import {
  parseFile,
  exportWorkbook,
  downloadBlob,
  SpreadsheetError,
} from '../services/spreadsheetService';
import FileDropzone from '../components/spreadsheet/FileDropzone';
import SpreadsheetToolbar from '../components/spreadsheet/SpreadsheetToolbar';
import SpreadsheetGrid from '../components/spreadsheet/SpreadsheetGrid';
import SheetTabs from '../components/spreadsheet/SheetTabs';
import WorkbookStatus from '../components/spreadsheet/WorkbookStatus';
import Icon from '../components/Icon';
import './Spreadsheet.css';

const Spreadsheet = () => {
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [selectedCell, setSelectedCell] = useState<CellAddress | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [diagnostics, setDiagnostics] = useState<{
    inputMounted: boolean;
    selectedFile: File | null;
    onChangeFired: boolean;
    processingStarted: boolean;
    processingCompleted: boolean;
    processingError: string | null;
  }>({
    inputMounted: false,
    selectedFile: null,
    onChangeFired: false,
    processingStarted: false,
    processingCompleted: false,
    processingError: null,
  });

  const handleFileSelect = useCallback(async (file: File) => {
    console.log('[UPLOAD] handleFileSelect called, file:', file?.name, file?.size, file?.type);
    setDiagnostics(prev => ({ ...prev, processingStarted: true, processingError: null }));
    try {
      setError(null);
      const wb = await parseFile(file);
      console.log('[UPLOAD] parseFile resolved, workbook sheets:', wb?.sheets?.length);
      setWorkbook(wb);
      setSelectedCell({ row: 0, col: 0 });
      setDiagnostics(prev => ({ ...prev, processingCompleted: true }));
      console.log('[UPLOAD] React state update requested (setWorkbook)');
    } catch (err) {
      console.error('[UPLOAD ERROR] handleFileSelect:', err);
      const msg = err instanceof SpreadsheetError ? err.message : 'An unexpected error occurred while reading the file.';
      setError(msg);
      setDiagnostics(prev => ({ ...prev, processingError: msg }));
    }
  }, []);

  const handleExport = useCallback(async () => {
    if (!workbook) return;
    try {
      setIsExporting(true);
      setError(null);
      const blob = await exportWorkbook(workbook, {
        format: 'xlsx',
        fileName: workbook.metadata.fileName,
      });
      downloadBlob(blob, workbook.metadata.fileName.replace(/\.[^/.]+$/, '') + '_modified.xlsx');
    } catch (err) {
      if (err instanceof SpreadsheetError) {
        setError(err.message);
      } else {
        setError('Failed to export the workbook.');
      }
    } finally {
      setIsExporting(false);
    }
  }, [workbook]);

  const handleExportCsv = useCallback(async () => {
    if (!workbook) return;
    try {
      setIsExporting(true);
      setError(null);
      const activeSheet = workbook.sheets[workbook.activeSheetIndex];
      const blob = await exportWorkbook(workbook, {
        format: 'csv',
        sheetName: activeSheet?.name,
      });
      const baseName = workbook.metadata.fileName.replace(/\.[^/.]+$/, '');
      downloadBlob(blob, `${baseName}_${activeSheet?.name || 'sheet'}.csv`);
    } catch (err) {
      if (err instanceof SpreadsheetError) {
        setError(err.message);
      } else {
        setError('Failed to export the workbook.');
      }
    } finally {
      setIsExporting(false);
    }
  }, [workbook]);

  const handleSheetSelect = useCallback((index: number) => {
    if (!workbook) return;
    setWorkbook((prev) =>
      prev
        ? {
            ...prev,
            activeSheetIndex: index,
          }
        : prev,
    );
    // Reset selection to first cell of new sheet
    const ws = workbook.sheets[index];
    if (ws && ws.rowCount > 0 && ws.colCount > 0) {
      setSelectedCell({ row: 0, col: 0 });
    }
  }, [workbook]);

  const handleCellSelect = useCallback((address: CellAddress) => {
    setSelectedCell(address);
  }, []);

  const handleCellEdit = useCallback((address: CellAddress, value: CellValue) => {
    if (!workbook) return;

    const activeSheet = workbook.sheets[workbook.activeSheetIndex];
    if (!activeSheet) return;

    // Create a new workbook with updated cell data
    setWorkbook((prev) => {
      if (!prev) return prev;

      const newSheets = prev.sheets.map((ws, wsIndex) => {
        if (wsIndex !== prev.activeSheetIndex) return ws;

        const newRows = ws.rows.map((row) => ({
          ...row,
          cells: row.cells.map((c) => {
            if (c.address.row === address.row && c.address.col === address.col) {
              return {
                ...c,
                value,
                modified: c.original !== value,
              };
            }
            return c;
          }),
        }));

        return {
          ...ws,
          rows: newRows,
        };
      });

      return {
        ...prev,
        sheets: newSheets,
      };
    });
  }, [workbook]);

  // Get the currently active worksheet
  const activeWorksheet: Worksheet | undefined = workbook
    ? workbook.sheets[workbook.activeSheetIndex]
    : undefined;

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    console.log('[UPLOAD] handleFileInputChange fired, files count:', e.target?.files?.length || 0);
    const files = e.target.files;
    if (files && files.length > 0) {
      console.log('[UPLOAD] File object:', files[0]?.name, files[0]?.size, files[0]?.type);
      setDiagnostics(prev => ({
        ...prev,
        onChangeFired: true,
        selectedFile: files[0],
      }));
      handleFileSelect(files[0]);
      e.target.value = '';
    } else {
      console.log('[UPLOAD] No files selected');
    }
  };

  // Render empty state
  if (!workbook) {
    return (
      <div className="spreadsheet-empty">
        <div className="spreadsheet-empty__content">
          <div className="spreadsheet-empty__icon" aria-hidden="true">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
              <path d="M14 2v6h6" />
              <path d="M8 13h8" />
              <path d="M8 17h8" />
              <path d="M12 9v8" />
            </svg>
          </div>
          <h2 className="spreadsheet-empty__title">No spreadsheet loaded</h2>
          <p className="spreadsheet-empty__description">
            Upload an Excel (.xlsx, .xls) or CSV (.csv) file to start editing.
            You can also drag and drop files directly onto this area.
          </p>
          <p className="spreadsheet-empty__formats">Supported: .xlsx, .xls, .csv (max 10 MB)</p>

          {/* DIAGNOSTICS: Visible native file input — Phase 1 */}
          <div style={{ margin: '1.5rem 0', padding: '1rem', backgroundColor: '#f0f0f0', borderRadius: '8px', border: '2px solid #3b82f6' }}>
            <p style={{ fontSize: '0.75rem', color: '#666', marginBottom: '0.5rem' }}>
              <strong>Upload Diagnostic</strong> — Native input test
            </p>
            <p style={{ fontSize: '0.75rem', color: '#333' }}>
              Input mounted: <strong>{diagnostics.inputMounted ? 'YES' : 'NO'}</strong> |
              onChange fired: <strong>{diagnostics.onChangeFired ? 'YES' : 'NO'}</strong> |
              Selected file: <strong>{diagnostics.selectedFile?.name || 'None'}</strong>
            </p>
            <p style={{ fontSize: '0.75rem', color: '#333' }}>
              Processing: <strong>{diagnostics.processingStarted ? 'STARTED' : 'NOT STARTED'}</strong> |
              Completed: <strong>{diagnostics.processingCompleted ? 'YES' : 'NO'}</strong>
              {diagnostics.processingError && <span style={{ color: 'red' }}> | Error: {diagnostics.processingError}</span>}
            </p>
          </div>

          {/* NATIVE VISIBLE INPUT — Phase 1 diagnostic */}
          <input
            ref={fileInputRef}
            id="spreadsheet-file-input-empty"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileInputChange}
            style={{
              display: 'block',
              padding: '0.5rem',
              border: '2px solid #3b82f6',
              borderRadius: '0.375rem',
              backgroundColor: '#fff',
              cursor: 'pointer',
              marginBottom: '1rem',
              fontSize: '0.875rem',
              width: '100%',
              maxWidth: '300px',
            }}
            aria-label="Select spreadsheet file (native input)"
          />

          <FileDropzone
            onFileSelect={handleFileSelect}
            accept=".xlsx,.xls,.csv"
            maxSizeMB={10}
          />
          {error && (
            <div className="spreadsheet-empty__error" style={{ color: '#ef4444', marginTop: '1rem', padding: '0.75rem', backgroundColor: '#fef2f2', border: '1px solid #fecaca', borderRadius: '0.375rem', fontSize: '0.875rem' }}>
              Upload failed: {error}
            </div>
          )}
        </div>
      </div>
    );
  }

  // Render workspace with loaded workbook
  return (
    <div className="spreadsheet-page">
      <input
        ref={fileInputRef}
        id="spreadsheet-file-input"
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFileInputChange}
        className="sr-only"
        aria-label="Select spreadsheet file"
        tabIndex={-1}
      />

      <SpreadsheetToolbar
        fileName={workbook.metadata.fileName}
        activeSheetName={activeWorksheet?.name || '—'}
        selectedCell={selectedCell}
        uploadInputId="spreadsheet-file-input"
        onExport={handleExport}
        onExportCsv={handleExportCsv}
        onFileInputChange={handleFileInputChange}
      />

      <WorkbookStatus workbook={workbook} error={null} onClearError={() => setError(null)} />

      {error && (
        <div className="spreadsheet-status spreadsheet-status--error" role="alert">
          <Icon name="alertCircle" size={16} />
          <span>{error}</span>
        </div>
      )}
      {!isExporting && (
        <div className="spreadsheet-workspace">
          <SheetTabs
            sheets={workbook.sheets}
            activeSheetIndex={workbook.activeSheetIndex}
            onSheetSelect={handleSheetSelect}
          />

          {activeWorksheet && (
            <SpreadsheetGrid
              worksheet={activeWorksheet}
              selectedCell={selectedCell}
              onCellSelect={handleCellSelect}
              onCellEdit={handleCellEdit}
            />
          )}
        </div>
      )}
    </div>
  );
};

export default Spreadsheet;