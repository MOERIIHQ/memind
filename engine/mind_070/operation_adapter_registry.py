"""Phase B registry: binds each allowed operation to its adapter callable.

This module satisfies the frozen spec (MIND-070 Specification v1.0, 3.2):
- the registry explicitly defines the allowed operations, their input/output
  contracts, and their adapter, and
- it does NOT contain engine execution logic.

The five Phase B adapters live in engine.mind_070.adapters. This module only
binds those callables into an OperationRegistry.
"""

from __future__ import annotations

from engine.mind_070.operation_registry import OperationType
from engine.mind_070.operation_registry import (
    OperationContract,
    OperationInfo,
    OperationRegistry,
)
from engine.mind_070.adapters import (
    adapt_auto_reporting,
    adapt_file_enhance,
    adapt_formula_apply,
    adapt_spreadsheet_intelligence,
    adapt_text_task_to_formula,
)

__all__ = ["build_phase_b_registry", "get_default_registry"]

_ADAPTERS = {
    "text_task_to_formula": adapt_text_task_to_formula,
    "formula_apply": adapt_formula_apply,
    "file_enhance": adapt_file_enhance,
    "spreadsheet_intelligence": adapt_spreadsheet_intelligence,
    "auto_reporting": adapt_auto_reporting,
}

_CONTRACTS = {
    "text_task_to_formula": OperationContract(
        input={"text": "str", "input_path": "str", "output_path": "str"},
        output={"status": "str", "output_path": "str", "formula_applied": "str"},
        engine="engine.text_task_parser",
        requires_ai=False,
        side_effects=["read_input", "write_output", "validate"],
    ),
    "formula_apply": OperationContract(
        input={"input_path": "str", "output_path": "str", "operation": "dict"},
        output={"status": "str", "output_path": "str", "formula_persisted": "bool"},
        engine="engine.formula_engine",
        requires_ai=False,
        side_effects=["read_input", "write_output", "validate"],
    ),
    "file_enhance": OperationContract(
        input={"input_path": "str", "output_path": "str"},
        output={"status": "str", "output_path": "str", "issues_found": "int", "valid": "bool"},
        engine="engine.file_enhancer",
        requires_ai=False,
        side_effects=["read_input", "write_output", "validate"],
    ),
    "spreadsheet_intelligence": OperationContract(
        input={"input_path": "str"},
        output={"status": "str", "spreadsheet_analysis": "dict"},
        engine="engine.mind_050.spreadsheet_intelligence",
        requires_ai=False,
        side_effects=["read_input"],
    ),
    "auto_reporting": OperationContract(
        input={"spreadsheet_analysis": "dict"},
        output={"status": "str", "report": "dict"},
        engine="engine.mind_060.auto_reporting",
        requires_ai=False,
        side_effects=["build_output_only"],
    ),
}


def _info(name: str) -> OperationInfo:
    return OperationInfo(
        name=name,
        operation_type=OperationType(name.upper().replace("-", "_").replace(" ", "_")),
        contract=_CONTRACTS[name],
        adapter=_ADAPTERS[name],
        enabled=True,
    )


def build_phase_b_registry() -> OperationRegistry:
    """Return a fresh registry bound to the five Phase B adapters."""
    registry = OperationRegistry()
    for name in _ADAPTERS.keys():
        registry.register_operation(_info(name))
    return registry


# Shared, module-level registry (tests may build their own via build_phase_b_registry).
_default_registry: OperationRegistry = None


def get_default_registry() -> OperationRegistry:
    global _default_registry
    if _default_registry is None:
        _default_registry = build_phase_b_registry()
    return _default_registry
