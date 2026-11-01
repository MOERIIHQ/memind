"""Phase B engine adapters.

Each adapter function maps one workflow step (with its parameters) onto the real
MIND-010..060 public API, returning a StepResult the executor records.

Contract rules:
- Only public methods of the MIND components are used. There is no new openpyxl
  access outside MIND-010 (engine/excel_engine.py); engines only use their own
  public API and do not duplicate spreadsheet machinery.
- Engine-specific exceptions are translated into the MIND-070 unified error model
  (WorkflowError / StepResult) and surfaced in diagnostics; they are never
  swallowed.
- Side effects are limited to: read the input workbook, write the output
  workbook where the component writes, and read-only analysis.
"""

from __future__ import annotations

from pathlib import Path

from engine.mind_070.contracts import (
    Step,
    StepResult,
    StepStatus,
    WorkflowError,
    WorkflowErrorCategory,
)
from engine.mind_070.operation_registry import OperationType

from engine.text_task_parser import TaskParseError, parse as parse_text, TextTaskParser
from engine.formula_engine import FormulaEngine, FormulaOperation, FormulaEngineError
from engine.file_enhancer import FileEnhancer, FileEnhancerError
from engine.mind_050.spreadsheet_intelligence import SpreadsheetIntelligence
from engine.mind_060.auto_reporting import generate_report, Report

__all__ = [
    "adapt_text_task_to_formula",
    "adapt_formula_apply",
    "adapt_file_enhance",
    "adapt_spreadsheet_intelligence",
    "adapt_auto_reporting",
]


# ---------------------------------------------------------------- helpers ---

def _failed(
    step: Step,
    message: str,
    category: WorkflowErrorCategory,
    *,
    diagnostics: dict = None,
    retryable: bool = False,
) -> StepResult:
    """Build a FAILED StepResult for an invalid step."""
    return StepResult(
        status=StepStatus.FAILED,
        output={"error": message},
        artifact=None,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"error": message, "category": category.value, **(diagnostics or {})},
    )


def _result_from_exception(step: Step, exc: Exception) -> StepResult:
    """Map an unexpected engine exception onto a FAILED StepResult (Phase B minimal).

    Phase C carries the full unified-error + provenance model.
    """
    category = _category_from_exception(exc)
    return StepResult(
        status=StepStatus.FAILED,
        output={"error": str(exc), "code": getattr(exc, "code", None)},
        artifact=None,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"error": str(exc), "category": category.value},
    )


def _category_from_exception(exc: Exception) -> WorkflowErrorCategory:
    """Classify engine exceptions for MIND-070 error categories (Phase B)."""
    if isinstance(exc, TaskParseError):
        return WorkflowErrorCategory.VALIDATION_FAILED
    if isinstance(exc, (WorkflowError,)):
        return exc.category
    # MIND-010/020/030/050 exceptions: execution failure.
    if isinstance(exc, (FileEnhancerError, FormulaEngineError, ValueError, OSError)):
        return WorkflowErrorCategory.EXECUTION_FAILURE
    return WorkflowErrorCategory.EXECUTION_FAILURE


# --------------------------------------------------------------- adapters ---

def adapt_text_task_to_formula(context, step: Step) -> StepResult:
    """MIND-040 parse + MIND-030 formula apply (delegated by TextTaskParser).

    Input parameters: text, input_path, output_path.
    """
    params = step.parameters
    if not isinstance(params, dict):
        return _failed(step, "step parameters must be a dict", WorkflowErrorCategory.VALIDATION_FAILED)
    text = params.get("text")
    input_path = params.get("input_path")
    output_path = params.get("output_path")
    if not text:
        return _failed(
            step,
            "missing 'text' parameter",
            WorkflowErrorCategory.VALIDATION_FAILED,
            diagnostics={"kind": "INVALID", "reason": "empty_input"},
            retryable=False,
        )
    if not input_path or not output_path:
        return _failed(step, "missing input_path or output_path", WorkflowErrorCategory.VALIDATION_FAILED)

    parsed = parse_text(text)
    if isinstance(parsed, TaskParseError):
        return _failed(
            step,
            parsed.message,
            WorkflowErrorCategory.VALIDATION_FAILED,
            diagnostics={"kind": parsed.kind, "details": parsed.details},
            retryable=False,
        )

    try:
        parser = TextTaskParser()
        engine_result = parser.execute(parsed, input_path, output_path)
    except Exception as exc:  # pragma: no cover - defensive
        return _result_from_exception(step, exc)

    return StepResult(
        status=StepStatus.SUCCEEDED,
        output=engine_result,
        artifact=output_path,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"engine": "mind_040", "component": "TextTaskParser"},
    )


def adapt_formula_apply(context, step: Step) -> StepResult:
    """MIND-030 apply a formula.

    step.parameters: input_path, output_path, operation={target, operation,
    operands, overwrite_policy, sheet}.
    """
    params = step.parameters
    if not isinstance(params, dict):
        return _failed(step, "step parameters must be a dict", WorkflowErrorCategory.VALIDATION_FAILED)
    input_path = params.get("input_path")
    output_path = params.get("output_path")
    op_payload = params.get("operation")
    if not input_path or not output_path or not op_payload:
        return _failed(
            step,
            "missing input_path, output_path, or operation",
            WorkflowErrorCategory.VALIDATION_FAILED,
        )
    try:
        formula_op = FormulaOperation(
            target=str(op_payload.get("target", "")),
            operation=str(op_payload.get("operation", "")),
            operands=op_payload.get("operands", {}) or {},
            overwrite_policy=str(op_payload.get("overwrite_policy", "never")),
            sheet=op_payload.get("sheet"),
        )
    except (AttributeError, TypeError) as exc:
        return _failed(step, f"invalid operation contract: {exc}", WorkflowErrorCategory.VALIDATION_FAILED)

    try:
        engine = FormulaEngine()
        try:
            engine_result = engine.apply_formula(input_path, output_path, formula_op)
        finally:
            engine.close()
    except Exception as exc:  # pragma: no cover - defensive
        return _result_from_exception(step, exc)

    return StepResult(
        status=StepStatus.SUCCEEDED,
        output=engine_result,
        artifact=output_path,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"engine": "mind_030", "component": "FormulaEngine"},
    )


def adapt_file_enhance(context, step: Step) -> StepResult:
    """MIND-020 FileEnhancer.enhance."""
    params = step.parameters
    if not isinstance(params, dict):
        return _failed(step, "step parameters must be a dict", WorkflowErrorCategory.VALIDATION_FAILED)
    input_path = params.get("input_path")
    output_path = params.get("output_path")
    if not input_path or not output_path:
        return _failed(step, "missing input_path or output_path", WorkflowErrorCategory.VALIDATION_FAILED)
    try:
        enhancer = FileEnhancer()
        result = enhancer.enhance(input_path, output_path)
    except Exception as exc:  # pragma: no cover - defensive
        return _result_from_exception(step, exc)

    return StepResult(
        status=StepStatus.SUCCEEDED,
        output={
            "original_sheet_count": result.original_sheet_count,
            "enhanced_sheet_count": result.enhanced_sheet_count,
            "issues_found": result.issues_found,
            "enhancements_applied": result.enhancements_applied,
            "cells_trimmed": result.cells_trimmed,
            "valid": result.valid,
        },
        artifact=output_path,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"engine": "mind_020", "component": "FileEnhancer"},
    )


def adapt_spreadsheet_intelligence(context, step: Step) -> StepResult:
    """MIND-050 SpreadsheetIntelligence.analyze (read-only)."""
    params = step.parameters
    if not isinstance(params, dict):
        return _failed(step, "step parameters must be a dict", WorkflowErrorCategory.VALIDATION_FAILED)
    input_path = params.get("input_path")
    if not input_path:
        return _failed(step, "missing input_path", WorkflowErrorCategory.VALIDATION_FAILED)
    try:
        engine = SpreadsheetIntelligence()
        try:
            result = engine.analyze(Path(input_path))
        finally:
            engine.close()
    except Exception as exc:  # pragma: no cover - defensive
        return _result_from_exception(step, exc)

    return StepResult(
        status=StepStatus.SUCCEEDED,
        output=result,
        artifact=None,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"engine": "mind_050", "component": "SpreadsheetIntelligence"},
    )


def adapt_auto_reporting(context, step: Step) -> StepResult:
    """MIND-060 generate_report(analysis).

    Input parameters: spreadsheet_analysis (SpreadsheetAnalysis instance,
    produced by the MIND-050 adapter).
    """
    params = step.parameters
    if not isinstance(params, dict):
        return _failed(step, "step parameters must be a dict", WorkflowErrorCategory.VALIDATION_FAILED)
    analysis = params.get("spreadsheet_analysis")
    if analysis is None:
        return _failed(step, "missing required 'spreadsheet_analysis'", WorkflowErrorCategory.VALIDATION_FAILED)
    try:
        report = generate_report(analysis)
    except (AttributeError, TypeError) as exc:
        return _failed(
            step,
            f"spreadsheet_analysis must be a SpreadsheetAnalysis instance, got {type(analysis).__name__}: {exc}",
            WorkflowErrorCategory.VALIDATION_FAILED,
        )
    except Exception as exc:  # pragma: no cover - defensive
        return _result_from_exception(step, exc)

    return StepResult(
        status=StepStatus.SUCCEEDED,
        output=report.to_dict(),
        artifact=None,
        retry_attempt=0,
        max_retries=1,
        diagnostics={"engine": "mind_060", "component": "AutoReporting"},
    )
