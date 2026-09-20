/**
 * MeMind Core Excel Engine — Spreadsheet Service
 *
 * Handles all file parsing and export logic, wrapping SheetJS (xlsx).
 * This is the single source of truth for converting between:
 *   - Raw File/Blob → Workbook (in-memory data model)
 *   - Workbook → File (exported Blob)
 */

import * as XLSX from 'xlsx';
import type {
  Cell,
  CellAddress,
  CellType,
  CellValue,
  CellAddressString,
  Workbook,
  WorkbookMetadata,
  Worksheet,
  WorksheetColumn,
  WorksheetRow,
  CellModification,
} from '../types/spreadsheet';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Supported file extensions and MIME types */
const SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.csv'];
const SUPPORTED_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'text/csv',
  'application/csv',
  'text/anytext',
];

// ---------------------------------------------------------------------------
// Interfaces
// ---------------------------------------------------------------------------

export interface ParseResult {
  workbook: Workbook;
  modifications: CellModification[];
}

export interface ExportOptions {
  format: 'xlsx' | 'csv';
  fileName?: string;
  sheetName?: string;
}

// ---------------------------------------------------------------------------
// Error class
// ---------------------------------------------------------------------------

export class SpreadsheetError extends Error {
  code: string;

  constructor(message: string, code: string) {
    super(message);
    this.name = 'SpreadsheetError';
    this.code = code;
  }
}

// ---------------------------------------------------------------------------
// File validation
// ---------------------------------------------------------------------------

/**
 * Validate that a File object is of a supported type.
 * Throws SpreadsheetError if invalid.
 */
export function validateFile(file: File): void {
  if (!file) {
    throw new SpreadsheetError('No file provided.', 'NO_FILE');
  }

  const ext = getFileExtension(file.name).toLowerCase();
  const isSupportedExt = SUPPORTED_EXTENSIONS.includes(ext);
  const isSupportedMime = SUPPORTED_MIME_TYPES.includes(file.type);

  console.log('[UPLOAD] validateFile:', { name: file.name, ext, extValid: isSupportedExt, fileType: file.type, mimeValid: isSupportedMime });

  if (!isSupportedExt && !isSupportedMime) {
    throw new SpreadsheetError(
      `Unsupported file type. Supported formats: ${SUPPORTED_EXTENSIONS.join(', ')}`,
      'UNSUPPORTED_TYPE',
    );
  }
}

/**
 * Extract the file extension from a filename (including the dot).
 */
export function getFileExtension(filename: string): string {
  const lastDot = filename.lastIndexOf('.');
  return lastDot >= 0 ? filename.slice(lastDot) : '';
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

/**
 * Read a File and parse it into a Workbook.
 * This is async because file reading + SheetJS parsing can be large.
 */
export async function parseFile(file: File): Promise<Workbook> {
  console.log('[UPLOAD] parseFile called, file:', file?.name, file?.size, file?.type);
  validateFile(file);

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        console.log('[UPLOAD] FileReader.onload fired');
        const data = e.target?.result as ArrayBuffer;
        console.log('[UPLOAD] ArrayBuffer created, byteLength:', data?.byteLength);
        const workbook = parseArrayBuffer(data, file);
        console.log('[UPLOAD] parseArrayBuffer succeeded');
        resolve(workbook);
      } catch (err) {
        console.error('[UPLOAD ERROR] parseFile onload:', err);
        reject(err);
      }
    };

    reader.onerror = () => {
      console.error('[UPLOAD ERROR] FileReader.onerror fired');
      reject(new SpreadsheetError('Failed to read the file. Please try again.', 'READ_ERROR'));
    };

    reader.readAsArrayBuffer(file);
    console.log('[UPLOAD] FileReader.readAsArrayBuffer called');
  });
}

/**
 * Convert an ArrayBuffer into a Workbook data model.
 */
export function parseArrayBuffer(data: ArrayBuffer, file: File): Workbook {
  try {
    console.log('[UPLOAD] XLSX.read starting');
    const workbook = XLSX.read(data, {
      type: 'array',
      cellDates: true,
      cellNF: true,
      cellFormula: true,
      cellText: false,
    });
    console.log('[UPLOAD] XLSX.read completed, SheetNames:', workbook?.SheetNames);

    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new SpreadsheetError('This workbook contains no sheets.', 'EMPTY_WORKBOOK');
    }

    const sheets: Worksheet[] = workbook.SheetNames.map((sheetName, position) => {
      const rawSheet = workbook.Sheets[sheetName] || {};
      return parseWorksheet(rawSheet, sheetName, position);
    });

    // Validate that at least one sheet has data
    const anySheetHasData = sheets.some((s) => s.rowCount > 0 && s.colCount > 0);
    if (!anySheetHasData) {
      throw new SpreadsheetError('No usable data found in any sheet.', 'NO_DATA');
    }

    const metadata: WorkbookMetadata = {
      fileName: file.name,
      fileSize: file.size,
      fileType: getFileExtension(file.name),
      loadedAt: new Date(),
      author: (workbook.Props?.Author as string) || null,
      title: (workbook.Props?.Title as string) || null,
    };

    return {
      metadata,
      sheets,
      activeSheetIndex: 0,
      _rawWorkbook: workbook,
    };
  } catch (err) {
    console.error('[UPLOAD ERROR] parseArrayBuffer error:', err);
    if (err instanceof SpreadsheetError) throw err;
    throw new SpreadsheetError(
      'Failed to parse the spreadsheet. The file may be corrupted or in an unexpected format.',
      'PARSE_ERROR',
    );
  }
}

/**
 * Parse a single SheetJS sheet object into our Worksheet data model.
 */
function parseWorksheet(
  rawSheet: XLSX.WorkSheet,
  name: string,
  position: number,
): Worksheet {
  // Use SheetJS range to determine dimensions
  const ref = rawSheet['!ref'] || 'A1';
  const range = XLSX.utils.decode_range(ref);
  const rowCount = range.e.r - range.s.r + 1;
  const colCount = range.e.c - range.s.c + 1;

  // Read cell data as 2D array (raw values)
  const jsonData = XLSX.utils.sheet_to_json(rawSheet, {
    header: 1,
    defval: null,
    raw: true,
  }) as unknown[][];

  const rows: WorksheetRow[] = [];
  const columns: WorksheetColumn[] = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const cells: Cell[] = [];
    for (let colIndex = 0; colIndex < colCount; colIndex++) {
      const cellRef = XLSX.utils.encode_cell({ c: colIndex, r: rowIndex });
      const rawCell = rawSheet[cellRef];
      const jsonValue = jsonData[rowIndex]?.[colIndex] ?? null;
      cells.push(buildCell(rawCell, jsonValue, colIndex, rowIndex));
    }
    rows.push({
      index: rowIndex,
      cells,
      height: getRowHeight(rawSheet, rowIndex),
    });
  }

  // Build column metadata
  for (let colIndex = 0; colIndex < colCount; colIndex++) {
    columns.push({
      index: colIndex,
      key: colToLetterStatic(colIndex),
      width: getColumnWidth(rawSheet, colIndex),
      visible: colIndex >= 0,
    });
  }

  return {
    name,
    position,
    rows,
    columns,
    rowCount,
    colCount,
    protected: isSheetProtected(rawSheet),
    _rawSheet: rawSheet,
  };
}

/**
 * Build a Cell object from a raw SheetJS cell.
 */
function buildCell(
  rawCell: XLSX.CellObject | undefined,
  jsonValue: unknown,
  col: number,
  row: number,
): Cell {
  const address: CellAddress = { row, col };
  const addrStr = `${colToLetterStatic(col)}${row + 1}` as CellAddressString;

  if (!rawCell) {
    return {
      value: null,
      original: null,
      type: 'z',
      formula: null,
      address,
      modified: false,
      addressString: addrStr,
    };
  }

  // Determine cell type from the raw cell type character
  const cellType: CellType = (rawCell.t as CellType) || 's';

  // Determine the value to display/edit
  let displayValue: CellValue;
  if (rawCell.f) {
    // Cell has a formula — preserve formula, use cached value for display
    displayValue = jsonValueToCell(jsonValue);
  } else {
    displayValue = jsonValueToCell(jsonValue);
  }

  return {
    value: displayValue,
    original: displayValue,
    type: cellType,
    formula: rawCell.f || null,
    address,
    modified: false,
    addressString: addrStr,
  };
}

/**
 * Convert a JSON-extracted value into our CellValue type.
 */
function jsonValueToCell(jsonValue: unknown): CellValue {
  if (jsonValue === null || jsonValue === undefined) return null;
  if (jsonValue instanceof Date) return jsonValue;
  if (typeof jsonValue === 'number') return jsonValue;
  if (typeof jsonValue === 'boolean') return jsonValue;
  if (typeof jsonValue === 'string') {
    const trimmed = jsonValue.trim();
    if (trimmed === '') return null;
    if (/^-?\d+\.?\d*$/.test(trimmed)) {
      const num = parseFloat(trimmed);
      if (!isNaN(num)) return num;
    }
    return jsonValue;
  }
  return String(jsonValue);
}

/**
 * Convert a CellValue back to a string representation for display.
 */
export function cellValueToString(value: CellValue): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().split('T')[0]; // YYYY-MM-DD
  if (typeof value === 'number') return formatNumber(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

/**
 * Format a number for display (trim trailing zeros for integers).
 */
function formatNumber(num: number): string {
  if (Number.isInteger(num)) return String(num);
  return String(parseFloat(num.toFixed(10)));
}

// ---------------------------------------------------------------------------
// Row/Column helpers
// ---------------------------------------------------------------------------

function colToLetterStatic(col: number): string {
  let letter = '';
  let tmp = col;
  while (tmp >= 0) {
    letter = String.fromCharCode((tmp % 26) + 65) + letter;
    tmp = Math.floor(tmp / 26) - 1;
  }
  return letter;
}

function getRowHeight(sheet: XLSX.WorkSheet, row: number): number | null {
  const rowObj = sheet['!rows']?.[row];
  if (rowObj && rowObj.hpx) return rowObj.hpx;
  if (rowObj && rowObj.hpt) return Math.round((rowObj.hpt * 3) / 4); // pt to px
  return null;
}

function getColumnWidth(sheet: XLSX.WorkSheet, col: number): number | null {
  const colObj = sheet['!cols']?.[col];
  if (colObj && colObj.wpx) return colObj.wpx;
  if (colObj && colObj.wch) return colObj.wch * 7; // approximate
  return null;
}

function isSheetProtected(sheet: XLSX.WorkSheet): boolean {
  return !!(sheet as Record<string, unknown>)['!protect'];
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

/**
 * Export the current workbook (with any in-memory modifications) to a Blob.
 * This rebuilds a SheetJS workbook from the data model, preserving formulas.
 */
export async function exportWorkbook(
  workbook: Workbook,
  options: ExportOptions = { format: 'xlsx' },
): Promise<Blob> {
  try {
    const output: XLSX.WorkBook = XLSX.utils.book_new();

    for (const ws of workbook.sheets) {
      const sheet = worksheetToSheetJS(ws);
      XLSX.utils.book_append_sheet(output, sheet, ws.name);
    }

    // Copy workbook properties if available
    const rawWb = workbook._rawWorkbook as XLSX.WorkBook | undefined;
    if (rawWb && rawWb.Props) {
      output.Props = rawWb.Props;
    }

    // Generate the file
    if (options.format === 'csv') {
      const sheetName = options.sheetName || workbook.sheets[workbook.activeSheetIndex]?.name;
      if (!sheetName) {
        throw new SpreadsheetError('No sheet available to export as CSV.', 'NO_SHEET');
      }
      const activeSheet = workbook.sheets.find((s: Worksheet) => s.name === sheetName);
      if (!activeSheet) {
        throw new SpreadsheetError(`Sheet "${sheetName}" not found.`, 'NO_SHEET');
      }
      const csv = XLSX.utils.sheet_to_csv(worksheetToSheetJS(activeSheet));
      return new Blob([csv], { type: 'text/csv' });
    }

    const wbout = XLSX.write(output, {
      bookType: 'xlsx',
      type: 'array',
      compression: true,
    });

    return new Blob([wbout], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  } catch (err) {
    if (err instanceof SpreadsheetError) throw err;
    throw new SpreadsheetError('Failed to export the workbook. Please try again.', 'EXPORT_ERROR');
  }
}

/**
 * Convert our Worksheet data model back to a SheetJS sheet object.
 * Preserves formulas by writing the formula string into the cell's 'f' field.
 */
function worksheetToSheetJS(ws: Worksheet): XLSX.WorkSheet {
  const data: unknown[][] = [];

  for (let r = 0; r < ws.rowCount; r++) {
    const row: unknown[] = [];
    for (let c = 0; c < ws.colCount; c++) {
      const cell = ws.rows[r]?.cells[c];
      if (cell && cell.value !== null) {
        row.push(cell.value);
      } else {
        row.push(undefined);
      }
    }
    data.push(row);
  }

  const sheet = XLSX.utils.aoa_to_sheet(data);

  // Apply formulas — overwrite cell objects to preserve formula on export
  for (let r = 0; r < ws.rowCount; r++) {
    for (let c = 0; c < ws.colCount; c++) {
      const cell = ws.rows[r]?.cells[c];
      if (cell && cell.formula) {
        const cellRef = XLSX.utils.encode_cell({ c, r });
        // Create or update the cell object to include the formula
        if (!sheet[cellRef]) {
          sheet[cellRef] = XLSX.utils.aoa_to_sheet([[cell.value]])[XLSX.utils.encode_cell({ c: 0, r: 0 })];
        }
        sheet[cellRef].f = cell.formula;
        sheet[cellRef].t = cell.type;
      }
    }
  }

  // Apply column widths if available
  const colSizes: { wch?: number }[] = [];
  for (let c = 0; c < ws.colCount; c++) {
    const col = ws.columns[c];
    if (col && col.width) {
      colSizes[c] = { wch: Math.max(8, Math.round(col.width / 7)) };
    }
  }
  if (colSizes.length > 0 && colSizes.some((c) => c.wch !== undefined)) {
    sheet['!cols'] = colSizes;
  }

  return sheet;
}

/**
 * Trigger a browser download of a Blob.
 */
export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// Re-export utilities
// ---------------------------------------------------------------------------

export { colToLetterStatic as colToLetter };
