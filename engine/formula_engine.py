"""
MeMind Formula Engine — MIND-030

Deterministic Excel formula engine providing:
- Formula detection (per-cell metadata)
- Structural validation within declared scope
- Structured formula generation
- Explicit formula application
- Existing formula protection
- Protected-sheet safety
- Save/reopen verification

Architecture:
    FormulaEngine  →  MIND-010 CoreExcelEngine (public API only)  →  .xlsx

Does NOT implement:
- AI/LLM formula generation
- Natural-language formula interpretation
- Full Excel calculation engine
- Complete Excel formula parser
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional

from engine.excel_engine import CoreExcelEngine, ExcelError, ExcelErrorCode


# ---------------------------------------------------------------------------
# Error types
# ---------------------------------------------------------------------------

class FormulaEngineError(Exception):
    """Structured error for the Formula Engine."""

    def __init__(self, message: str, code: str, details: Optional[dict] = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}


class FormulaEngineErrorCode(Enum):
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    INVALID_FORMAT = "INVALID_FORMAT"
    WORKBOOK_CORRUPT = "WORKBOOK_CORRUPT"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    INVALID_CELL_REF = "INVALID_CELL_REF"
    INVALID_OPERATION = "INVALID_OPERATION"
    UNSUPPORTED_OPERATION = "UNSUPPORTED_OPERATION"
    FORMULA_INVALID = "FORMULA_INVALID"
    SAVE_FAILED = "SAVE_FAILED"
    VALIDATION_FAILED = "VALIDATION_FAILED"
    PROTECTED_SHEET = "PROTECTED_SHEET"
    EXISTING_FORMULA_PROTECTED = "EXISTING_FORMULA_PROTECTED"


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class FormulaType(Enum):
    """Broad classification of formula structure."""
    ARITHMETIC = "arithmetic"        # =A1+B1, =A1*B1
    FUNCTION = "function"            # =SUM(A1:A10), =IF(...)
    COMPARISON = "comparison"        # =A1>B1, =A1<=B1
    REFERENCE = "reference"          # =A1, =Sheet1!A1
    UNKNOWN = "unknown"              # Cannot classify


@dataclass
class FormulaInfo:
    """Metadata about a detected formula cell."""
    sheet: str
    cell: str              # Address like "A1"
    row: int               # 0-based row index
    col: int               # 0-based column index
    formula: str           # Full formula string including "="
    formula_type: str      # FormulaType value
    references: list[str] = field(default_factory=list)  # Extracted cell refs


@dataclass
class FormulaValidationResult:
    """Result of structural formula validation."""
    valid: bool
    sheet: str
    cell: str
    formula: str
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return self.valid


@dataclass
class FormulaOperation:
    """Structured instruction for formula generation and application."""
    target: str                        # Cell address like "D2"
    operation: str                     # SUM, AVERAGE, ADD, SUBTRACT, MULTIPLY, DIVIDE, IF, MIN, MAX, COUNT, COUNTA
    operands: dict                     # Operation-specific operands
    overwrite_policy: str = "never"   # "never" or "explicit"
    sheet: Optional[str] = None       # Target sheet; defaults to first sheet


# ---------------------------------------------------------------------------
# Formula Engine
# ---------------------------------------------------------------------------

class FormulaEngine:
    """
    MIND-030 Deterministic Formula Engine.

    Uses only public MIND-010 API.
    """

    # Supported operations for structured generation
    SUPPORTED_OPERATIONS = {
        "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA",
        "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE",
        "IF",
    }

    # Supported function names (for validation)
    KNOWN_FUNCTIONS = {
        "SUM", "AVERAGE", "AVERAGEA", "MIN", "MAX", "COUNT",
        "COUNTA", "COUNTBLANK", "IF", "AND", "OR", "NOT",
        "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "SQRT",
        "LEN", "CONCATENATE", "TEXTJOIN",
    }

    # Cell reference pattern: A1, $A$1, A$1, $A1, A1:B2, Sheet1!A1, Sheet1!A1:B2
    CELL_REF_PATTERN = re.compile(
        r'(?:'
        r'([A-Za-z_][A-Za-z0-9_]*)!?'  # optional sheet name
        r'?)?'
        r'\$?[A-Z]+\$?\d+'
        r'(?:'
        r':\$?[A-Z]+\$?\d+'
        r')?'
    )

    def __init__(self):
        self._engine = CoreExcelEngine()

    # -----------------------------------------------------------------------
    # LOAD / LIFECYCLE
    # -----------------------------------------------------------------------

    def load(self, file_path: str) -> None:
        """Load a workbook for formula operations."""
        try:
            self._engine.load(file_path)
        except ExcelError as e:
            raise FormulaEngineError(
                f"Cannot load workbook: {e.message}",
                e.code,
                {"original_error": e.code},
            ) from e

    def close(self) -> None:
        """Close the workbook."""
        self._engine.close()

    # -----------------------------------------------------------------------
    # A. INSPECT / DETECT
    # -----------------------------------------------------------------------

    def scan_formulas(self, file_path: str) -> list[FormulaInfo]:
        """
        Detect all formula cells in a workbook.

        Returns list of FormulaInfo for each formula cell.
        Does NOT modify the workbook.
        Does NOT evaluate formulas.

        Pipeline: LOAD → INSPECT → DETECT → CLOSE
        """
        self.load(file_path)
        result = self._scan_formulas_loaded()
        self.close()
        return result

    def _scan_formulas_loaded(self) -> list[FormulaInfo]:
        """Scan formula cells in the already-loaded workbook."""
        info = self._engine.get_workbook_info()
        sheet_names = self._engine.get_sheet_names()
        results: list[FormulaInfo] = []

        for sheet_name in sheet_names:
            sheet_data = self._engine.get_sheet(sheet_name)
            max_row = sheet_data.row_count or 1
            max_col = sheet_data.col_count or 1

            for r in range(max_row):
                for c in range(max_col):
                    if r < len(sheet_data.cells) and c < len(sheet_data.cells[r]):
                        cell = sheet_data.cells[r][c]
                        val = cell.value
                        if isinstance(val, str) and val.startswith("="):
                            addr = self._row_col_to_address(r, c)
                            ftype, refs = self._classify_formula(val)
                            results.append(FormulaInfo(
                                sheet=sheet_name,
                                cell=addr,
                                row=r,
                                col=c,
                                formula=val,
                                formula_type=ftype,
                                references=refs,
                            ))

        return results

    def get_formula(self, sheet_name: str, cell_address: str) -> Optional[FormulaInfo]:
        """
        Get detailed info for a specific formula cell.

        Uses get_cell() (public API) which populates CellData.formula.
        Returns None if cell has no formula.
        """
        row, col = self._address_to_row_col(cell_address)
        cell = self._engine.get_cell(sheet_name, row + 1, col + 1)
        if cell.formula:
            ftype, refs = self._classify_formula(cell.formula)
            return FormulaInfo(
                sheet=sheet_name,
                cell=cell_address,
                row=row,
                col=col,
                formula=cell.formula,
                formula_type=ftype,
                references=refs,
            )
        return None

    # -----------------------------------------------------------------------
    # B. VALIDATION
    # -----------------------------------------------------------------------

    def validate_formula(self, formula: str, sheet: str = "", cell: str = "") -> FormulaValidationResult:
        """
        Structurally validate a formula string.

        Checks:
        - Starts with "="
        - Balanced parentheses
        - Known function names (if function call detected)
        - Basic cell reference syntax

        Does NOT calculate or verify computed results.
        """
        errors: list[str] = []
        warnings: list[str] = []

        if not formula:
            errors.append("Formula is empty")
            return FormulaValidationResult(False, sheet, cell, formula, errors, warnings)

        if not formula.startswith("="):
            errors.append("Formula must start with '='")
            return FormulaValidationResult(False, sheet, cell, formula, errors, warnings)

        formula_body = formula[1:]  # Strip leading "="

        # Check balanced parentheses
        paren_depth = 0
        for ch in formula_body:
            if ch == '(':
                paren_depth += 1
            elif ch == ')':
                paren_depth -= 1
            if paren_depth < 0:
                errors.append("Unbalanced parentheses: extra closing ')'")
                break
        if paren_depth > 0:
            errors.append(f"Unbalanced parentheses: {paren_depth} unclosed '('")

        # Check for unknown/malformed function calls
        func_matches = re.findall(r'([A-Za-z_][A-Za-z0-9_]*)\s*\(', formula_body)
        for func_name in func_matches:
            if func_name.upper() not in self.KNOWN_FUNCTIONS:
                errors.append(f"Unsupported function: {func_name}() — outside MIND-030 MVP scope")

        # Check for references to potentially invalid sheet names (heuristic only)
        sheet_refs = re.findall(r"([A-Za-z_][A-Za-z0-9_]*)!\$?[A-Z]+\$?\d+", formula_body)
        # We cannot verify sheet existence without loading a workbook — report as info

        # Check for dangling operators at end
        stripped = formula_body.strip()
        if stripped and stripped[-1] in '+-*/=':
            errors.append(f"Dangling operator at end: '{stripped[-1]}'")

        # Check for empty function arguments
        if re.search(r'\(\s*\)\s*$', formula_body):
            errors.append("Empty function arguments")

        # Check for LET/LAMBDA/advanced function patterns even without parens (edge case)
        advanced_patterns = ['LET', 'LAMBDA', 'XLOOKUP', 'XIRR', 'XMATCH', 'FILTER', 'SORT', 'UNIQUE', 'SEQUENCE']
        for pat in advanced_patterns:
            if re.search(rf'\b{pat}\b', formula_body):
                # Only flag if not already caught as function call above
                if not any(func_name.upper() == pat for func_name in func_matches):
                    errors.append(f"Unsupported function: {pat}() — outside MIND-030 MVP scope")

        valid = len(errors) == 0
        return FormulaValidationResult(valid, sheet, cell, formula, errors, warnings)

    # -----------------------------------------------------------------------
    # C. GENERATION
    # -----------------------------------------------------------------------

    def generate_formula(self, operation: FormulaOperation) -> str:
        """
        Generate a formula string from a structured operation.

        Raises FormulaEngineError for unsupported operations or invalid operands.

        Examples:
            {"target": "D2", "operation": "SUM", "operands": {"range": "B2:C2"}}
            → "=SUM(B2:C2)"

            {"target": "F2", "operation": "MULTIPLY", "operands": {"left": "D2", "right": "E2"}}
            → "=D2*E2"
        """
        op = operation.operation.upper()
        operands = operation.operands

        if op not in self.SUPPORTED_OPERATIONS:
            raise FormulaEngineError(
                f"Unsupported operation: {op}. Supported: {', '.join(sorted(self.SUPPORTED_OPERATIONS))}",
                FormulaEngineErrorCode.UNSUPPORTED_OPERATION.value,
                {"operation": op},
            )

        if op == "SUM":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "SUM requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=SUM({rng})"

        elif op == "AVERAGE":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "AVERAGE requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=AVERAGE({rng})"

        elif op == "MIN":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "MIN requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=MIN({rng})"

        elif op == "MAX":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "MAX requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=MAX({rng})"

        elif op == "COUNT":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "COUNT requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=COUNT({rng})"

        elif op == "COUNTA":
            rng = operands.get("range")
            if not rng:
                raise FormulaEngineError(
                    "COUNTA requires 'range' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f"=COUNTA({rng})"

        elif op == "ADD":
            left = operands.get("left")
            right = operands.get("right")
            if not left or not right:
                raise FormulaEngineError(
                    "ADD requires 'left' and 'right' operands",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            self._validate_cell_ref(left)
            self._validate_cell_ref(right)
            return f"={left}+{right}"

        elif op == "SUBTRACT":
            left = operands.get("left")
            right = operands.get("right")
            if not left or not right:
                raise FormulaEngineError(
                    "SUBTRACT requires 'left' and 'right' operands",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            self._validate_cell_ref(left)
            self._validate_cell_ref(right)
            return f"={left}-{right}"

        elif op == "MULTIPLY":
            left = operands.get("left")
            right = operands.get("right")
            if not left or not right:
                raise FormulaEngineError(
                    "MULTIPLY requires 'left' and 'right' operands",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            self._validate_cell_ref(left)
            self._validate_cell_ref(right)
            return f"={left}*{right}"

        elif op == "DIVIDE":
            left = operands.get("left")
            right = operands.get("right")
            if not left or not right:
                raise FormulaEngineError(
                    "DIVIDE requires 'left' and 'right' operands",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            self._validate_cell_ref(left)
            self._validate_cell_ref(right)
            return f"={left}/{right}"

        elif op == "IF":
            condition = operands.get("condition")
            true_val = operands.get("true_val")
            false_val = operands.get("false_val")
            if condition is None:
                raise FormulaEngineError(
                    "IF requires 'condition' operand",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            # true_val and false_val are optional in Excel IF, but we require them for clarity
            if true_val is None or false_val is None:
                raise FormulaEngineError(
                    "IF requires 'true_val' and 'false_val' operands",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            return f'=IF({condition},"{true_val}","{false_val}")'

        # Fallback (should not reach here)
        raise FormulaEngineError(
            f"Unhandled operation: {op}",
            FormulaEngineErrorCode.UNSUPPORTED_OPERATION.value,
        )

    # -----------------------------------------------------------------------
    # D. APPLICATION
    # -----------------------------------------------------------------------

    def apply_formula(
        self,
        input_path: str,
        output_path: str,
        operation: FormulaOperation,
    ) -> dict[str, Any]:
        """
        Apply a formula to a target cell in the workbook.

        Safety rules:
        1. Existing formulas are NOT overwritten by default (overwrite_policy="never").
        2. Explicit overwrite requires overwrite_policy="explicit".
        3. Invalid target cell → rejected.
        4. Missing sheet → rejected.
        5. Invalid operation → rejected.
        6. Protected sheet → rejected.
        7. Workbook remains saveable.
        8. Formula verified after reopen.

        Pipeline: LOAD → VALIDATE → GENERATE → CHECK SAFETY → APPLY → SAVE → REOPEN → VERIFY
        """
        # Load
        try:
            self._engine.load(input_path)
        except ExcelError as e:
            raise FormulaEngineError(
                f"Cannot load workbook: {e.message}",
                e.code,
            ) from e

        # Determine target sheet
        target_sheet = operation.sheet
        if target_sheet is None:
            sheet_names = self._engine.get_sheet_names()
            if not sheet_names:
                raise FormulaEngineError(
                    "Workbook has no sheets",
                    FormulaEngineErrorCode.INVALID_OPERATION.value,
                )
            target_sheet = sheet_names[0]

        # Validate sheet exists
        if target_sheet not in self._engine.get_sheet_names():
            raise FormulaEngineError(
                f"Sheet not found: {target_sheet}",
                FormulaEngineErrorCode.SHEET_NOT_FOUND.value,
                {"sheet": target_sheet},
            )

        # Check protected sheet
        if self._engine.is_sheet_protected(target_sheet):
            raise FormulaEngineError(
                f"Cannot apply formula to protected sheet: {target_sheet}",
                FormulaEngineErrorCode.PROTECTED_SHEET.value,
                {"sheet": target_sheet},
            )

        # Parse target address
        target_row, target_col = self._address_to_row_col(operation.target)
        if target_row is None or target_col is None:
            raise FormulaEngineError(
                f"Invalid target cell address: {operation.target}",
                FormulaEngineErrorCode.INVALID_CELL_REF.value,
                {"target": operation.target},
            )

        # Validate target cell is within sheet bounds
        sheet_data = self._engine.get_sheet(target_sheet)
        if target_row >= sheet_data.row_count or target_col >= sheet_data.col_count:
            # Row/col beyond existing data — this is allowed in Excel (you can write to any cell)
            # But we check for completely invalid references (row/col < 0)
            pass

        if target_row < 0 or target_col < 0:
            raise FormulaEngineError(
                f"Invalid cell reference: row={target_row}, col={target_col}",
                FormulaEngineErrorCode.INVALID_CELL_REF.value,
            )

        # Generate the formula
        try:
            formula = self.generate_formula(operation)
        except FormulaEngineError:
            raise
        except Exception as e:
            raise FormulaEngineError(
                f"Formula generation failed: {e}",
                FormulaEngineErrorCode.INVALID_OPERATION.value,
            ) from e

        # Check existing cell content
        existing_cell = self._engine.get_cell(target_sheet, target_row + 1, target_col + 1)
        has_existing_formula = (
            existing_cell.formula is not None
            and isinstance(existing_cell.value, str)
            and existing_cell.value.startswith("=")
        )

        if has_existing_formula and operation.overwrite_policy != "explicit":
            raise FormulaEngineError(
                f"Cell {operation.target} already contains a formula. "
                f"Set overwrite_policy='explicit' to replace.",
                FormulaEngineErrorCode.EXISTING_FORMULA_PROTECTED.value,
                {"cell": operation.target, "existing_formula": existing_cell.value},
            )

        # Apply the formula
        self._engine.update_cell(target_sheet, target_row + 1, target_col + 1, formula)

        # Save
        try:
            saved_path = self._engine.save(output_path)
        except ExcelError as e:
            raise FormulaEngineError(
                f"Failed to save workbook: {e.message}",
                e.code,
            ) from e

        # Validate
        try:
            validation = self._engine.validate(saved_path)
        except ExcelError as e:
            raise FormulaEngineError(
                f"Validation failed: {e.message}",
                e.code,
            ) from e

        self._engine.close()

        # Verify formula persistence by reopening
        verify_engine = CoreExcelEngine()
        try:
            verify_engine.load(saved_path)
            verified_cell = verify_engine.get_cell(target_sheet, target_row + 1, target_col + 1)
            formula_persisted = (
                verified_cell.formula is not None
                and verified_cell.formula == formula
            )
            verify_engine.close()
        except Exception as e:
            formula_persisted = False
            if verify_engine._wb is not None:
                verify_engine.close()

        result: dict[str, Any] = {
            "valid": validation.get("valid", False),
            "output_path": saved_path,
            "formula_applied": formula,
            "target_cell": operation.target,
            "target_sheet": target_sheet,
            "formula_persisted": formula_persisted,
            "overwrite_allowed": operation.overwrite_policy == "explicit",
            "errors": validation.get("errors", []),
        }

        return result

    # -----------------------------------------------------------------------
    # BATCH APPLY
    # -----------------------------------------------------------------------

    def apply_formulas(
        self,
        input_path: str,
        output_path: str,
        operations: list[FormulaOperation],
    ) -> list[dict[str, Any]]:
        """
        Apply multiple formulas in sequence.

        Each operation is applied independently.
        If one fails, previous operations are NOT rolled back
        (the workbook is saved at each step).

        Use with caution — for transactional safety, apply one at a time.
        """
        results: list[dict[str, Any]] = []
        current_input = input_path

        for i, op in enumerate(operations):
            temp_output = output_path if i == len(operations) - 1 else f"{output_path}.tmp.{i}"
            try:
                result = self.apply_formula(current_input, temp_output, op)
                results.append(result)
                current_input = temp_output
            except FormulaEngineError as e:
                results.append({
                    "valid": False,
                    "error": e.message,
                    "code": e.code,
                    "operation_index": i,
                    "target": op.target,
                })
                # Continue with next operation using original input
                # (no partial state corruption)
                current_input = input_path

        # If all succeeded, ensure final output is at output_path
        if all(r.get("valid", False) for r in results):
            if current_input != output_path and os.path.exists(current_input):
                import shutil
                shutil.move(current_input, output_path)
                results[-1]["output_path"] = output_path

        return results

    # -----------------------------------------------------------------------
    # VALIDATION HELPERS
    # -----------------------------------------------------------------------

    def validate_cell_reference(self, ref: str) -> bool:
        """Check if a string looks like a valid cell reference (basic syntax check)."""
        return bool(self.CELL_REF_PATTERN.fullmatch(ref))

    def _validate_cell_ref(self, ref: str) -> None:
        """Raise if ref doesn't look like a valid cell reference."""
        if not self.validate_cell_reference(ref):
            raise FormulaEngineError(
                f"Invalid cell reference: {ref}",
                FormulaEngineErrorCode.INVALID_CELL_REF.value,
                {"reference": ref},
            )

    # -----------------------------------------------------------------------
    # PRIVATE HELPERS
    # -----------------------------------------------------------------------

    def _classify_formula(self, formula: str) -> tuple[str, list[str]]:
        """
        Classify a formula and extract cell references.
        Returns (formula_type, references).
        """
        formula_body = formula[1:]  # Strip "="
        refs: list[str] = []

        # Extract all cell references from the formula
        for match in self.CELL_REF_PATTERN.finditer(formula_body):
            ref = match.group(0)
            if ref not in refs:
                refs.append(ref)

        # Classify
        if re.match(r'^[A-Za-z_][A-Za-z0-9_]*\s*\(', formula_body):
            return (FormulaType.FUNCTION.value, refs)
        elif any(ch in formula_body for ch in '<>=') and not re.match(r'^[A-Za-z_][A-Za-z0-9_]*\s*\(', formula_body):
            # Contains comparison operators but isn't a function
            if '>' in formula_body or '<' in formula_body or '<>' in formula_body:
                return (FormulaType.COMPARISON.value, refs)
        elif any(ch in formula_body for ch in '+-*/'):
            return (FormulaType.ARITHMETIC.value, refs)
        elif refs:
            return (FormulaType.REFERENCE.value, refs)
        else:
            return (FormulaType.UNKNOWN.value, refs)

    def _address_to_row_col(self, address: str) -> tuple[Optional[int], Optional[int]]:
        """
        Convert cell address like 'A1' or 'B2' to (row, col) 0-based.
        Returns (None, None) if invalid.
        """
        address = address.strip()
        if not address:
            return (None, None)

        # Match pattern: optional sheet! then column letters then row number
        match = re.match(r'(?:([A-Za-z_][A-Za-z0-9_]*)!)?(\$?[A-Z]+)(\$?\d+)', address)
        if not match:
            return (None, None)

        col_letters = match.group(2).replace('$', '')
        row_num = int(match.group(3).replace('$', ''))

        # Convert column letters to index
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
        col -= 1  # 0-based

        row = row_num - 1  # 0-based

        if row < 0 or col < 0:
            return (None, None)

        return (row, col)

    def _row_col_to_address(self, row: int, col: int) -> str:
        """Convert 0-based row/col to Excel address like 'A1'."""
        from openpyxl.utils import get_column_letter
        return f"{get_column_letter(col + 1)}{row + 1}"

    def _get_sheet_or_error(self, sheet_name: str):
        """Use public API to get sheet or raise."""
        return self._engine._get_sheet_or_error(sheet_name)
