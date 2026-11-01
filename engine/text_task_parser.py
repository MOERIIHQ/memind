"""
MeMind Text → Task Parser — MIND-040

Deterministic text command parser that converts limited spreadsheet
instructions into structured tasks for the Formula Engine (MIND-030).

Architecture:
    Text → MIND-040 Parser → TextTask | TaskParseError
                                     ↓
                            FormulaEngine.apply_formula()

MIND-040 is NOT:
- a general chatbot
- an LLM system
- a general NLP system
- a calculation engine
- a workflow automation engine

Supported MVP grammar:
- SUM: "jumlahkan <range> ke <target>" / "jumlahkan <start> sampai <end> ke <target>"
- MULTIPLY: "kalikan <left> dengan <right> ke <target>"

Author: Hermes Agent (executor)
License: Same as MeMind project
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from engine.formula_engine import (
    FormulaEngine,
    FormulaOperation,
    FormulaEngineError,
)


# ---------------------------------------------------------------------------
# Parse error model
# ---------------------------------------------------------------------------

@dataclass
class TaskParseError:
    """Explicit parse-level error for MIND-040.

    Kinds:
    - "AMBIGUOUS": required information cannot be deterministically resolved
    - "INVALID": command matches grammar pattern but is malformed
    - "UNSUPPORTED": command is outside approved MIND-040 grammar
    """

    kind: str  # "AMBIGUOUS" | "INVALID" | "UNSUPPORTED"
    message: str
    details: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Structured task model
# ---------------------------------------------------------------------------

@dataclass
class TextTask:
    """Structured task produced by MIND-040 parser.

    A successfully parsed TextTask is inherently parse-valid.
    No mutable status field — success means parsed.
    """

    task_type: str = "FORMULA"
    operation: str = "SUM"  # "SUM" | "MULTIPLY"
    target: str = ""
    operands: dict = field(default_factory=dict)
    sheet: Optional[str] = None
    overwrite_policy: str = "never"


# ---------------------------------------------------------------------------
# Cell reference validation (reuse MIND-030 contract)
# ---------------------------------------------------------------------------

# Reuse the same pattern as FormulaEngine for consistency
_CELL_REF_PATTERN = re.compile(
    r'(?:'
    r'([A-Za-z_][A-Za-z0-9_]*)!?'  # optional sheet name
    r'?)?'
    r'\$?[A-Z]+\$?\d+'
    r'(?:'
    r':\$?[A-Z]+\$?\d+'
    r')?'
)

_RANGE_PATTERN = re.compile(
    r'\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+'
)


def validate_cell_reference(ref: str) -> bool:
    """Validate a cell reference against MIND-030 contract.

    Reuses the same regex pattern as FormulaEngine for consistency.
    """
    return bool(_CELL_REF_PATTERN.fullmatch(ref))


def validate_range(range_str: str) -> bool:
    """Validate a range reference like B2:C2."""
    return bool(_RANGE_PATTERN.fullmatch(range_str))


def extract_cell_ref(text: str) -> Optional[str]:
    """Extract a single cell reference from text.

    Returns the first valid cell reference found, or None.
    """
    match = _CELL_REF_PATTERN.search(text)
    if match:
        return match.group(0)
    return None


def extract_range(text: str) -> Optional[str]:
    """Extract a range reference (e.g., B2:C2) from text."""
    match = _RANGE_PATTERN.search(text)
    if match:
        return match.group(0)
    return None


def normalize_cell_ref(ref: str) -> str:
    """Strip $ signs to get base cell reference.

    Excel accepts both $D$2 and D2 for the same cell.
    For MVP consistency, normalize to base form without $.
    """
    return ref.replace('$', '')


def normalize_range(range_str: str) -> str:
    """Normalize range by stripping $ signs."""
    return range_str.replace('$', '')


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

class TextTaskParser:
    """Deterministic, stateless text → task parser for MIND-040.

    Supported grammar:
    - SUM: "jumlahkan <range> ke <target>"
           "jumlahkan <start> sampai <end> ke <target>"
    - MULTIPLY: "kalikan <left> dengan <right> ke <target>"

    The parser is:
    - Deterministic: same input → same output
    - Stateless: no mutable state between parses
    - Side-effect free: no file access, no workbook mutation
    - Safe: no eval/exec/subprocess/dynamic imports
    """

    # Operation whitelist — anything outside this is UNSUPPORTED
    SUPPORTED_OPERATIONS = {"SUM", "MULTIPLY"}

    def __init__(self):
        # No state — parser is stateless by design
        pass

    def parse(self, text: str) -> TextTask | TaskParseError:
        """Parse text into a TextTask or TaskParseError.

        Args:
            text: Human-readable spreadsheet command in Indonesian.

        Returns:
            TextTask if parsing succeeds, TaskParseError otherwise.
        """
        if not text or not text.strip():
            return TaskParseError(
                kind="INVALID",
                message="Command text is empty",
                details={"reason": "empty_input"},
            )

        text = text.strip()

        # Try SUM patterns first
        sum_result = self._try_parse_sum(text)
        if sum_result:
            return sum_result

        # Try MULTIPLY pattern
        multiply_result = self._try_parse_multiply(text)
        if multiply_result:
            return multiply_result

        # Check if text contains any recognizable spreadsheet intent
        # If it has keywords but doesn't match grammar, it's UNSUPPORTED
        if self._has_spreadsheet_intent(text):
            return TaskParseError(
                kind="UNSUPPORTED",
                message="Command is outside approved MIND-040 grammar",
                details={
                    "text": text,
                    "supported": ["jumlahkan ... ke ...", "kalikan ... dengan ... ke ..."],
                },
            )

        # No spreadsheet intent detected — INVALID
        return TaskParseError(
            kind="INVALID",
            message="Text does not match any supported command pattern",
            details={"text": text},
        )

    # -------------------------------------------------------------------
    # SUM pattern
    # -------------------------------------------------------------------

    def _try_parse_sum(self, text: str) -> TextTask | TaskParseError | None:
        """Try to parse SUM command.

        Supported forms:
        - "jumlahkan <range> ke <target>"
        - "jumlahkan <start> sampai <end> ke <target>"
        """
        if "jumlahkan" not in text.lower():
            return None

        # Find "ke " — target is after "ke"
        # Use pattern consistent with MIND-030 CELL_REF_PATTERN
        ke_match = re.search(r'ke\s+(\$?[A-Z]+\$?\d+)', text, re.IGNORECASE)
        if not ke_match:
            return TaskParseError(
                kind="INVALID",
                message="SUM command missing target after 'ke'",
                details={"text": text, "reason": "missing_target"},
            )

        target = ke_match.group(1)
        target = normalize_cell_ref(target)
        if not validate_cell_reference(target):
            return TaskParseError(
                kind="INVALID",
                message=f"Invalid target cell reference: {target}",
                details={"target": target, "reason": "invalid_reference"},
            )

        # Check for "sampai" pattern: "jumlahkan X sampai Y ke Z"
        if "sampai" in text.lower():
            # Extract range between "jumlahkan" and "sampai"
            start_idx = text.lower().find("jumlahkan") + len("jumlahkan")
            sampai_idx = text.lower().find("sampai")

            if start_idx >= sampai_idx:
                return TaskParseError(
                    kind="INVALID",
                    message="SUM command: 'sampai' must come after 'jumlahkan'",
                    details={"text": text, "reason": "bad_structure"},
                )

            # Extract what's between "jumlahkan" and "sampai"
            middle = text[start_idx:sampai_idx].strip()

            # Try to extract a range (X:Y) or two cell refs
            range_match = _RANGE_PATTERN.search(middle)
            if range_match:
                range_str = range_match.group(0)
                if not validate_range(range_str):
                    return TaskParseError(
                        kind="INVALID",
                        message=f"Invalid range reference: {range_str}",
                        details={"range": range_str, "reason": "invalid_range"},
                    )
                operands = {"range": normalize_range(range_str)}
            else:
                # Try to extract two cell references — from both sides of "sampai"
                # Left side: between "jumlahkan" and "sampai"
                # Right side: between "sampai" and "ke"
                refs = []
                for match in _CELL_REF_PATTERN.finditer(middle):
                    ref = match.group(0)
                    refs.append(ref)
                # Also check the text between "sampai" and "ke"
                # Find "ke" position
                ke_pos = ke_match.start()
                antara_sampai_ke = text[sampai_idx + len("sampai"):ke_pos].strip()
                for match in _CELL_REF_PATTERN.finditer(antara_sampai_ke):
                    ref = match.group(0)
                    if ref not in refs:  # Avoid duplicates
                        refs.append(ref)
                if len(refs) >= 2:
                    # Take first two references
                    left, right = refs[0], refs[1]
                    if not validate_cell_reference(left) or not validate_cell_reference(right):
                        return TaskParseError(
                            kind="INVALID",
                            message="Invalid cell references in SUM command",
                            details={"text": text, "reason": "invalid_references"},
                        )
                    # Construct range from two refs
                    operands = {"range": f"{normalize_cell_ref(left)}:{normalize_cell_ref(right)}"}
                else:
                    return TaskParseError(
                        kind="INVALID",
                        message="SUM with 'sampai' requires two cell references or a range",
                        details={"text": text, "reason": "missing_operands"},
                    )

            return TextTask(
                task_type="FORMULA",
                operation="SUM",
                target=target,
                operands=operands,
            )

        # Direct range form: "jumlahkan <range> ke <target>"
        # Look for range pattern anywhere before "ke"
        ke_pos = ke_match.start()
        before_ke = text[:ke_pos]

        range_match = _RANGE_PATTERN.search(before_ke)
        if range_match:
            range_str = range_match.group(0)
            if not validate_range(range_str):
                return TaskParseError(
                    kind="INVALID",
                    message=f"Invalid range reference: {range_str}",
                    details={"range": range_str, "reason": "invalid_range"},
                )
            return TextTask(
                task_type="FORMULA",
                operation="SUM",
                target=target,
                operands={"range": range_str},
            )

        # Try to find two cell refs (implicit range)
        refs = _CELL_REF_PATTERN.findall(before_ke)
        if len(refs) >= 2:
            left, right = refs[0], refs[1]
            if not validate_cell_reference(left) or not validate_cell_reference(right):
                return TaskParseError(
                    kind="INVALID",
                    message="Invalid cell references in SUM command",
                    details={"text": text, "reason": "invalid_references"},
                )
            # Filter out the target from refs if it appears
            if right == target:
                if len(refs) >= 3:
                    left, right = refs[0], refs[2]
                else:
                    return TaskParseError(
                        kind="AMBIGUOUS",
                        message="Cannot determine range — multiple interpretations possible",
                        details={"text": text, "reason": "ambiguous_range"},
                    )
            operands = {"range": f"{left}:{right}"}
            return TextTask(
                task_type="FORMULA",
                operation="SUM",
                target=target,
                operands=operands,
            )

        # No range or cell refs found
        return TaskParseError(
            kind="INVALID",
            message="SUM command requires a range or cell references",
            details={"text": text, "reason": "missing_range"},
        )

    # -------------------------------------------------------------------
    # MULTIPLY pattern
    # -------------------------------------------------------------------

    def _try_parse_multiply(self, text: str) -> TextTask | TaskParseError | None:
        """Try to parse MULTIPLY command.

        Supported form:
        - "kalikan <left> dengan <right> ke <target>"
        """
        if "kalikan" not in text.lower():
            return None

        # Find "ke " — target is after "ke"
        # Use pattern consistent with MIND-030 CELL_REF_PATTERN
        ke_match = re.search(r'ke\s+(\$?[A-Z]+\$?\d+)', text, re.IGNORECASE)
        if not ke_match:
            return TaskParseError(
                kind="INVALID",
                message="MULTIPLY command missing target after 'ke'",
                details={"text": text, "reason": "missing_target"},
            )

        target = ke_match.group(1)
        if not validate_cell_reference(target):
            return TaskParseError(
                kind="INVALID",
                message=f"Invalid target cell reference: {target}",
                details={"target": target, "reason": "invalid_reference"},
            )

        # Find "dengan" — separates left and right operands
        dengan_match = re.search(r'dengan\s+', text, re.IGNORECASE)
        if not dengan_match:
            return TaskParseError(
                kind="INVALID",
                message="MULTIPLY command missing 'dengan' separator between operands",
                details={"text": text, "reason": "missing_separator"},
            )

        # Extract text between "kalikan" and "dengan" → left operand
        kalikan_idx = text.lower().find("kalikan") + len("kalikan")
        dengan_idx = dengan_match.start()

        if kalikan_idx >= dengan_idx:
            return TaskParseError(
                kind="INVALID",
                message="'dengan' must come after 'kalikan'",
                details={"text": text, "reason": "bad_structure"},
            )

        left_text = text[kalikan_idx:dengan_idx].strip()
        left_refs = []
        for match in _CELL_REF_PATTERN.finditer(left_text):
            left_refs.append(match.group(0))
        if not left_refs:
            return TaskParseError(
                kind="INVALID",
                message="MULTIPLY command: left operand must be a cell reference",
                details={"text": text, "reason": "missing_left_operand"},
            )
        left = left_refs[0]
        left = normalize_cell_ref(left)
        if not validate_cell_reference(left):
            return TaskParseError(
                kind="INVALID",
                message=f"Invalid left cell reference: {left}",
                details={"left": left, "reason": "invalid_reference"},
            )

        # Extract text between "dengan" and "ke" → right operand
        dengan_end = dengan_match.end()
        ke_pos = ke_match.start()

        if dengan_end >= ke_pos:
            return TaskParseError(
                kind="INVALID",
                message="'ke' must come after 'dengan'",
                details={"text": text, "reason": "bad_structure"},
            )

        right_text = text[dengan_end:ke_pos].strip()
        right_refs = []
        for match in _CELL_REF_PATTERN.finditer(right_text):
            right_refs.append(match.group(0))
        if not right_refs:
            return TaskParseError(
                kind="INVALID",
                message="MULTIPLY command: right operand must be a cell reference",
                details={"text": text, "reason": "missing_right_operand"},
            )
        right = right_refs[0]
        right = normalize_cell_ref(right)
        if not validate_cell_reference(right):
            return TaskParseError(
                kind="INVALID",
                message=f"Invalid right cell reference: {right}",
                details={"right": right, "reason": "invalid_reference"},
            )

        return TextTask(
            task_type="FORMULA",
            operation="MULTIPLY",
            target=target,
            operands={"left": left, "right": right},
        )

    # -------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------

    def _has_spreadsheet_intent(self, text: str) -> bool:
        """Check if text contains keywords suggesting spreadsheet intent.

        Used to distinguish UNSUPPORTED from INVALID.
        """
        spreadsheet_keywords = [
            "jumlahkan", "kalikan", "tambahkan", "kurangkan",
            "hitung", "total", "rata", "min", "maks",
            "baris", "kolom", "sel", "cell", "sheet",
            "pivot", "grafik", "chart", "laporan",
        ]
        text_lower = text.lower()
        return any(kw in text_lower for kw in spreadsheet_keywords)

    # -------------------------------------------------------------------
    # Execute delegation
    # -------------------------------------------------------------------

    def execute(
        self,
        task: TextTask,
        input_path: str,
        output_path: str,
    ) -> dict:
        """Execute a parsed TextTask via MIND-030 FormulaEngine.

        This method delegates entirely to FormulaEngine.apply_formula().
        No direct CoreExcelEngine access.

        Args:
            task: Parsed TextTask
            input_path: Path to input XLSX
            output_path: Path for output XLSX

        Returns:
            Result dict from FormulaEngine.apply_formula()
        """
        # Construct FormulaOperation from TextTask
        operation = FormulaOperation(
            target=task.target,
            operation=task.operation,
            operands=task.operands,
            overwrite_policy=task.overwrite_policy,
            sheet=task.sheet,
        )

        # Delegate to MIND-030
        engine = FormulaEngine()
        try:
            result = engine.apply_formula(input_path, output_path, operation)
        finally:
            engine.close()

        return result


# ---------------------------------------------------------------------------
# Convenience function
# ---------------------------------------------------------------------------

def parse(text: str) -> TextTask | TaskParseError:
    """Convenience function for parsing text commands.

    Creates a parser instance and parses the text.
    Deterministic and stateless.
    """
    parser = TextTaskParser()
    return parser.parse(text)


def execute_task(
    task: TextTask,
    input_path: str,
    output_path: str,
) -> dict:
    """Convenience function for executing a parsed task.

    Delegates to FormulaEngine via TextTaskParser.execute().
    """
    parser = TextTaskParser()
    return parser.execute(task, input_path, output_path)
