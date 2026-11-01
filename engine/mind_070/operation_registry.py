"""Operation registry: allowed operations, data contracts, and adapters.

The registry is the single source of truth for which operations a workflow
plan may reference. It does not contain engine code. Operations are identified
by name and mapped to a concrete engine adapter during execution (Phase B).
Adapter names here are informational; they are resolved by the executor in
Phase B and are verified against the running codebase before binding.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable, Optional

__all__ = [
    "OperationType",
    "OperationContract",
    "OperationInfo",
    "OperationRegistry",
    "get_default_registry",
]


class OperationType(str, Enum):
    """Enumeration of operations the orchestrator can orchestrate."""

    TEXT_TASK_TO_FORMULA = "TEXT_TASK_TO_FORMULA"
    FORMULA_APPLY = "FORMULA_APPLY"
    FILE_ENHANCE = "FILE_ENHANCE"
    SPREADSHEET_INTELLIGENCE = "SPREADSHEET_INTELLIGENCE"
    AUTO_REPORTING = "AUTO_REPORTING"
    COMPOSITE = "COMPOSITE"


@dataclass(frozen=True)
class OperationContract:
    """Declared I/O contract of an operation."""

    input: dict[str, str] = field(default_factory=dict)
    output: dict[str, str] = field(default_factory=dict)
    engine: str = ""
    adapter: Optional[str] = None
    requires_ai: bool = False
    side_effects: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class OperationInfo:
    name: str
    operation_type: OperationType
    contract: OperationContract
    adapter: Optional[Callable[..., Any]] = None
    enabled: bool = True


class OperationRegistry:
    """Central registry of allowed operations and their contracts.

    Attributes
    ----------
    operations: names -> OperationInfo. Canonical source for plan validation.
    adapters: names -> callable. Bound in Phase B; the registry keeps them
        decoupled from the contract layer.
    """

    def __init__(self) -> None:
        self._operations: dict[str, OperationInfo] = {}
        self._adapters: dict[str, Callable[..., Any]] = {}
        self._register_defaults()

    # -- registry access --------------------------------------------------

    def register_operation(self, info: OperationInfo) -> OperationInfo:
        self._operations[info.name] = info
        return info

    def register_adapter(self, operation_name: str, adapter: Callable[..., Any]) -> None:
        self._adapters[operation_name] = adapter

    def get_operation(self, name: str) -> Optional[OperationInfo]:
        return self._operations.get(name)

    def is_available(self, name: str) -> bool:
        info = self._operations.get(name)
        return bool(info and info.enabled)

    def list_operations(self) -> list[str]:
        return [info.name for info in self._operations.values() if info.enabled]

    def list_adapters(self) -> list[str]:
        return list(self._adapters.keys())

    # -- defaults ---------------------------------------------------------

    def _register_defaults(self) -> None:
        self.register_operation(
            OperationInfo(
                name="text_task_to_formula",
                operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                contract=OperationContract(
                    input={
                        "text": "str",
                        "input_path": "str",
                        "output_path": "str",
                        "overwrite_policy": "str",
                    },
                    output={
                        "status": "str",
                        "output_path": "str",
                        "formula_applied": "str",
                    },
                    engine="engine.text_task_parser",
                    adapter="text_task_parser.TextTaskParser.execute",
                    requires_ai=False,
                    side_effects=["read_file", "write_file", "validate"],
                ),
                enabled=True,
            )
        )
        self.register_operation(
            OperationInfo(
                name="formula_apply",
                operation_type=OperationType.FORMULA_APPLY,
                contract=OperationContract(
                    input={
                        "target": "str",
                        "operation": "str",
                        "operands": "dict",
                        "input_path": "str",
                        "output_path": "str",
                        "overwrite_policy": "str",
                    },
                    output={
                        "status": "str",
                        "output_path": "str",
                        "formula_applied": "str",
                        "formula_persisted": "bool",
                    },
                    engine="engine.formula_engine",
                    adapter="formula_engine.FormulaEngine.apply_formula",
                    requires_ai=False,
                    side_effects=["read_file", "write_file", "validate"],
                ),
                enabled=True,
            )
        )
        self.register_operation(
            OperationInfo(
                name="file_enhance",
                operation_type=OperationType.FILE_ENHANCE,
                contract=OperationContract(
                    input={
                        "input_path": "str",
                        "output_path": "str",
                    },
                    output={
                        "status": "str",
                        "output_path": "str",
                        "issues_found": "int",
                        "enhancements_applied": "int",
                        "analysis": "dict",
                        "valid": "bool",
                    },
                    engine="engine.file_enhancer",
                    adapter="file_enhancer.FileEnhancer.enhance",
                    requires_ai=False,
                    side_effects=["read_file", "write_file", "validate"],
                ),
                enabled=True,
            )
        )
        self.register_operation(
            OperationInfo(
                name="spreadsheet_intelligence",
                operation_type=OperationType.SPREADSHEET_INTELLIGENCE,
                contract=OperationContract(
                    input={
                        "input_path": "str",
                    },
                    output={
                        "status": "str",
                        "spreadsheet_analysis": "dict",
                    },
                    engine="engine.mind_050.spreadsheet_intelligence",
                    adapter="spreadsheet_intelligence.SpreadsheetIntelligence.analyze",
                    requires_ai=False,
                    side_effects=["read_file", "analyze_only"],
                ),
                enabled=True,
            )
        )
        self.register_operation(
            OperationInfo(
                name="auto_reporting",
                operation_type=OperationType.AUTO_REPORTING,
                contract=OperationContract(
                    input={
                        "spreadsheet_analysis": "dict",
                        "output_path": "str",
                    },
                    output={
                        "status": "str",
                        "report_path": "str",
                        "report_sections": "list",
                    },
                    engine="engine.mind_060.auto_reporting",
                    adapter="auto_reporting.generate_report",
                    requires_ai=False,
                    side_effects=["build_output_only"],
                ),
                enabled=True,
            )
        )


# Module-level singleton used by workflow runners (tests can use a fresh
# instance via OperationRegistry()).
_default_registry: Optional[OperationRegistry] = None


def get_default_registry() -> OperationRegistry:
    global _default_registry
    if _default_registry is None:
        _default_registry = OperationRegistry()
    return _default_registry
