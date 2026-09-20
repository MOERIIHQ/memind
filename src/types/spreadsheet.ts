/**
 * MeMind Core Excel Engine — Data Model
 *
 * Designed to be extensible for future phases (Formula Engine, AI commands, etc.)
 * while remaining lightweight for this iteration.
 */

/**
 * A cell's data type, derived from what SheetJS gives us.
 */
export type CellType = 'n' | 's' | 'b' | 'd' | 'z' | 'e'; // number, string, boolean, date, zoned (stub), error

/**
 * The raw value of a cell — can be a number, string, boolean, Date, or null.
 */
export type CellValue = number | string | boolean | Date | null;

/**
 * Zero-based column index → letter(s). e.g. 0 → "A", 25 → "Z", 26 → "AA"
 */
export type CellAddress = {
  row: number;
  col: number;
};

export type CellAddressString = `${string}${number}`; // e.g. "A1", "BC3"

/**
 * Represents a single cell in the spreadsheet grid.
 * `value` is the user-visible / editable value.
 * `original` is the value as parsed from the source file (for diff tracking).
 * `formula` is preserved from the source if present (read-only this iteration).
 */
export interface Cell {
  /** Display/edit value. null = empty cell. */
  value: CellValue;
  /** Value as it was when first loaded from the file. */
  original: CellValue;
  /** Cell type as determined by the parser. */
  type: CellType;
  /** Formula string if the cell contains a formula (e.g. "SUM(A1:A2)"), else null. */
  formula: string | null;
  /** Zero-based row and column. */
  address: CellAddress;
  /** True if the user has edited this cell since loading. */
  modified: boolean;
  /** Human-readable address like "A1". Derived, not stored persistently. */
  addressString: CellAddressString;
}

/**
 * A single row in a worksheet.
 */
export interface WorksheetRow {
  /** Zero-based row index. */
  index: number;
  /** Cells in this row, keyed by zero-based column index. */
  cells: Cell[];
  /** Computed height in pixels (null = auto/default). */
  height: number | null;
}

/**
 * Column metadata for a worksheet.
 */
export interface WorksheetColumn {
  /** Zero-based column index. */
  index: number;
  /** Column letter(s), e.g. "A". */
  key: string;
  /** Width in characters (approximate). */
  width: number | null;
  /** Visible flag. */
  visible: boolean;
}

/**
 * A named sheet within a workbook.
 */
export interface Worksheet {
  /** Sheet name (tab label). */
  name: string;
  /** Zero-based position in the workbook. */
  position: number;
  /** Rows of cells. */
  rows: WorksheetRow[];
  /** Column metadata. */
  columns: WorksheetColumn[];
  /** Number of data rows. */
  rowCount: number;
  /** Number of data columns. */
  colCount: number;
  /** Sheet is protected. */
  protected: boolean;
  /** Raw SheetJS sheet object — preserved for formula/export fidelity. */
  _rawSheet: unknown;
}

/**
 * Metadata about the source file.
 */
export interface WorkbookMetadata {
  /** Original file name. */
  fileName: string;
  /** File size in bytes. */
  fileSize: number;
  /** MIME type or extension. */
  fileType: string;
  /** Date the file was parsed. */
  loadedAt: Date;
  /** Creator / author from workbook properties (if available). */
  author: string | null;
  /** Workbook title (if available). */
  title: string | null;
}

/**
 * The top-level in-memory representation of an entire workbook.
 */
export interface Workbook {
  /** Metadata about the loaded file. */
  metadata: WorkbookMetadata;
  /** All sheets in order. */
  sheets: Worksheet[];
  /** Index of the currently active sheet. */
  activeSheetIndex: number;
  /** Raw SheetJS workbook object — preserved for formula/export fidelity. */
  _rawWorkbook: unknown;
}

/**
 * A single cell modification tracked for export.
 */
export interface CellModification {
  sheetName: string;
  address: CellAddress;
  previousValue: CellValue;
  newValue: CellValue;
}

// ---------------------------------------------------------------------------
// Utility helpers (pure functions, no DOM dependencies)
// ---------------------------------------------------------------------------

/**
 * Convert a zero-based column index to a letter string (0 → "A", 26 → "AA").
 */
export function colToLetter(col: number): string {
  let letter = '';
  let tmp = col;
  while (tmp >= 0) {
    letter = String.fromCharCode((tmp % 26) + 65) + letter;
    tmp = Math.floor(tmp / 26) - 1;
  }
  return letter;
}

/**
 * Convert a letter string to a zero-based column index ("A" → 0, "AA" → 26).
 */
export function letterToCol(letter: string): number {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Convert a CellAddress to a human-readable address string ("A1").
 */
export function addressToString(addr: CellAddress): CellAddressString {
  return `${colToLetter(addr.col)}${addr.row + 1}` as CellAddressString;
}
