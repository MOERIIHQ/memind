export interface NavItem {
  id: string;
  label: string;
  icon: string;
  badge?: string | number;
  disabled?: boolean;
}

export interface FileItem {
  id: string;
  name: string;
  size: number;
  modified: Date;
  type: 'xlsx' | 'csv' | 'ods' | 'unknown';
  isMock?: boolean;
}

export interface QuickAction {
  id: string;
  label: string;
  description: string;
  icon: string;
  href?: string;
  onClick?: () => void;
  disabled?: boolean;
}

export interface ActivityItem {
  id: string;
  type: 'upload' | 'process' | 'analyze' | 'export' | 'error';
  title: string;
  description: string;
  timestamp: Date;
  status: 'pending' | 'processing' | 'completed' | 'failed';
}

export interface SpreadsheetData {
  id: string;
  name: string;
  sheets: SheetData[];
  createdAt: Date;
  updatedAt: Date;
}

export interface SheetData {
  name: string;
  rows: number;
  columns: number;
  data: (string | number | null)[][];
}

export interface PlaceholderPagePrimaryAction {
  label: string;
  primaryAction: () => void;
}

// Re-export spreadsheet types
export * from './spreadsheet';