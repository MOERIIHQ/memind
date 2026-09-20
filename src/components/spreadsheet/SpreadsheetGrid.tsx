import {
  useRef,
  useEffect,
  useCallback,
  type KeyboardEvent,
} from 'react';
import type { Cell, CellAddress, CellValue, Worksheet } from '../../types/spreadsheet';
import { cellValueToString } from '../../services/spreadsheetService';
import CellEditor from './CellEditor';
import './SpreadsheetGrid.css';

interface SpreadsheetGridProps {
  worksheet: Worksheet;
  selectedCell: CellAddress | null;
  onCellSelect: (address: CellAddress) => void;
  onCellEdit: (address: CellAddress, value: CellValue) => void;
}

const CELL_WIDTH = 100;

const SpreadsheetGrid = ({
  worksheet,
  selectedCell,
  onCellSelect,
  onCellEdit,
}: SpreadsheetGridProps) => {
  const gridRef = useRef<HTMLDivElement>(null);
  const cellRefs = useRef<Map<string, HTMLDivElement>>(new Map());
  const editingCellRef = useRef<CellAddress | null>(null);

  // Scroll to selected cell when it changes
  useEffect(() => {
    if (!selectedCell) return;
    const key = `${selectedCell.row}:${selectedCell.col}`;
    const el = cellRefs.current.get(key);
    if (el) {
      el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  }, [selectedCell]);

  // Focus the grid container for keyboard navigation
  const handleGridClick = () => {
    gridRef.current?.focus();
  };

  const getCellKey = (row: number, col: number) => `${row}:${col}`;

  const handleCellClick = (row: number, col: number) => {
    onCellSelect({ row, col });
    editingCellRef.current = null;
  };

  const startEditing = (row: number, col: number) => {
    onCellSelect({ row, col });
    editingCellRef.current = { row, col };
  };

  const getCell = (row: number, col: number): Cell | null => {
    if (row >= worksheet.rows.length) return null;
    const r = worksheet.rows[row];
    if (!r || col >= r.cells.length) return null;
    return r.cells[col];
  };

  const isCellSelected = (row: number, col: number): boolean => {
    if (!selectedCell) return false;
    return selectedCell.row === row && selectedCell.col === col;
  };

  const isCellEditing = (row: number, col: number): boolean => {
    const editing = editingCellRef.current;
    if (!editing) return false;
    return editing.row === row && editing.col === col;
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    if (!selectedCell) return;

    const { row, col } = selectedCell;
    let newRow = row;
    let newCol = col;

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        newRow = Math.max(0, row - 1);
        onCellSelect({ row: newRow, col });
        break;
      case 'ArrowDown':
        e.preventDefault();
        newRow = Math.min(worksheet.rowCount - 1, row + 1);
        onCellSelect({ row: newRow, col });
        break;
      case 'ArrowLeft':
        e.preventDefault();
        newCol = Math.max(0, col - 1);
        onCellSelect({ row, col: newCol });
        break;
      case 'ArrowRight':
        e.preventDefault();
        newCol = Math.min(worksheet.colCount - 1, col + 1);
        onCellSelect({ row, col: newCol });
        break;
      case 'Enter':
        e.preventDefault();
        startEditing(row, col);
        break;
      case ' ':
        e.preventDefault();
        startEditing(row, col);
        break;
      case 'Escape':
        e.preventDefault();
        editingCellRef.current = null;
        break;
      case 'Delete':
        e.preventDefault();
        commitCellEdit(row, col, null);
        break;
      case 'Backspace':
        e.preventDefault();
        commitCellEdit(row, col, null);
        break;
    }
  };

  const commitCellEdit = useCallback(
    (row: number, col: number, value: CellValue) => {
      onCellEdit({ row, col }, value);
      editingCellRef.current = null;
    },
    [onCellEdit],
  );

  const handleCellDoubleClick = (row: number, col: number) => {
    startEditing(row, col);
  };

  const handleEditorKeyDown = (_e: React.KeyboardEvent<HTMLInputElement>) => {
    // CellEditor handles Enter/Escape/Tab internally.
    // Grid-level navigation happens after commit via selectedCell changes.
  };

  const handleEditorCancel = () => {
    editingCellRef.current = null;
  };

  // Render cell content
  const renderCellContent = (cell: Cell | null) => {
    if (!cell || cell.value === null) return null;
    return cellValueToString(cell.value);
  };

  // Determine cell alignment based on value type
  const getCellClass = (cell: Cell | null): string => {
    if (!cell || cell.value === null) return '';
    if (typeof cell.value === 'number') return 'cell--numeric';
    if (typeof cell.value === 'boolean') return 'cell--boolean';
    return 'cell--text';
  };

  const setCellRef = useCallback((el: HTMLDivElement | null, key: string) => {
    if (el) cellRefs.current.set(key, el);
    else cellRefs.current.delete(key);
  }, []);

  // Render the grid body
  const renderGridBody = () => {
    const rows = [];
    for (let r = 0; r < worksheet.rowCount; r++) {
      const cells = [];
      for (let c = 0; c < worksheet.colCount; c++) {
        const cell = getCell(r, c);
        const isSelected = isCellSelected(r, c);
        const isEditing = isCellEditing(r, c);
        const cellKey = getCellKey(r, c);

        cells.push(
          <div
            key={cellKey}
            ref={(el) => setCellRef(el, cellKey)}
            className={`spreadsheet-grid__cell ${getCellClass(cell)} ${isSelected ? 'spreadsheet-grid__cell--selected' : ''}`}
            onClick={() => handleCellClick(r, c)}
            onDoubleClick={() => handleCellDoubleClick(r, c)}
            aria-colindex={c + 1}
            aria-rowindex={r + 1}
          >
            {isEditing && cell ? (
              <CellEditor
                cell={cell}
                isActive={true}
                onSubmit={(value) => commitCellEdit(r, c, value)}
                onCancel={handleEditorCancel}
                onKeyDown={handleEditorKeyDown}
              />
            ) : (
              <span className="spreadsheet-grid__cell-content">
                {renderCellContent(cell)}
              </span>
            )}
          </div>,
        );
      }
      rows.push(
        <div key={`row-${r}`} className="spreadsheet-grid__row" aria-rowindex={r + 1}>
          {/* Row number header */}
          <div className="spreadsheet-grid__row-header" aria-label={`Row ${r + 1}`}>
            {r + 1}
          </div>
          <div className="spreadsheet-grid__row-cells">
            {cells}
          </div>
        </div>,
      );
    }
    return rows;
  };

  return (
    <div
      ref={gridRef}
      className="spreadsheet-grid"
      tabIndex={-1}
      onClick={handleGridClick}
      onKeyDown={handleKeyDown}
      role="grid"
      aria-label="Spreadsheet grid"
    >
      {/* Column headers */}
      <div className="spreadsheet-grid__header-row">
        <div className="spreadsheet-grid__corner" aria-label="Corner"></div>
        {worksheet.columns.map((col) => (
          <div
            key={`col-header-${col.index}`}
            className="spreadsheet-grid__col-header"
            style={{ width: `${CELL_WIDTH}px`, minWidth: `${CELL_WIDTH}px` }}
            aria-label={`Column ${col.key}`}
          >
            {col.key}
          </div>
        ))}
      </div>

      {/* Body */}
      <div className="spreadsheet-grid__body">{renderGridBody()}</div>
    </div>
  );
};

export default SpreadsheetGrid;
