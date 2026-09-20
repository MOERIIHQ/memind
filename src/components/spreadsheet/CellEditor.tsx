import { useState, useRef, useEffect, type KeyboardEvent, type FocusEvent } from 'react';
import type { Cell, CellValue } from '../../types/spreadsheet';
import { cellValueToString } from '../../services/spreadsheetService';
import './CellEditor.css';

interface CellEditorProps {
  cell: Cell;
  isActive: boolean;
  onSubmit: (value: CellValue) => void;
  onCancel: () => void;
  onKeyDown: (e: KeyboardEvent<HTMLInputElement>) => void;
  onFocus?: () => void;
}

const CellEditor = ({
  cell,
  isActive,
  onSubmit,
  onCancel,
  onKeyDown,
  onFocus,
}: CellEditorProps) => {
  const [inputValue, setInputValue] = useState<string>(() => cellValueToString(cell.value));
  const inputRef = useRef<HTMLInputElement>(null);
  const initialValueRef = useRef<string>(cellValueToString(cell.value));

  // Focus the input when this editor becomes active
  useEffect(() => {
    if (isActive && inputRef.current) {
      inputRef.current.focus();
      const len = inputRef.current.value.length;
      inputRef.current.setSelectionRange(len, len);
    }
  }, [isActive]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setInputValue(e.target.value);
  };

  const handleBlur = (_e: FocusEvent<HTMLInputElement>) => {
    commitValue();
  };

  const commitValue = () => {
    const raw = inputValue.trim();
    let parsed: CellValue;

    if (raw === '') {
      parsed = null;
    } else if (/^-?\d+\.?\d*$/.test(raw)) {
      const num = parseFloat(raw);
      parsed = Number.isNaN(num) ? raw : num;
    } else if (raw.toLowerCase() === 'true') {
      parsed = true;
    } else if (raw.toLowerCase() === 'false') {
      parsed = false;
    } else {
      parsed = raw;
    }

    if (parsed !== cell.value) {
      onSubmit(parsed);
    } else {
      onCancel();
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    onKeyDown(e);

    if (e.key === 'Enter') {
      e.preventDefault();
      commitValue();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setInputValue(initialValueRef.current);
      onCancel();
    }
  };

  return (
    <input
      ref={inputRef}
      type="text"
      className="cell-editor"
      value={inputValue}
      onChange={handleChange}
      onKeyDown={handleKeyDown}
      onBlur={handleBlur}
      onFocus={onFocus}
      aria-label="Cell editor"
    />
  );
};

export default CellEditor;
