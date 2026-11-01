"""MIND-070A — shared contracts, statuses, and unified error model.

All contract types are JSON-serialization-friendly (dataclasses with plain
types) so they can be stored under ".memind/state/<workflow_id>/" without any
spreadsheet dependency. No engine code is imported here, so the contracts are
independent of MIND-010..MIND-060 implementation details.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Any, Optional


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


class WorkflowStatus(str, Enum):
    """Workflow-level execution status."""

    PLANNED = "PLANNED"
    RUNNING = "RUNNING"
    SUCCEEDED = "SUCCEEDED"
    FAILED = "FAILED"
    PARTIAL = "PARTIAL"
    CANCELLED = "CANCELLED"
    NEEDS_RECOVERY = "NEEDS_RECOVERY"


class StepStatus(str, Enum):
    """Per-step execution status."""

    PENDING = "PENDING"
    RUNNING = "RUNNING"
    SUCCEEDED = "SUCCEEDED"
    FAILED = "FAILED"
    PARTIAL = "PARTIAL"
    SKIPPED = "SKIPPED"
    CANCELLING = "CANCELLING"


class WorkflowErrorCategory(str, Enum):
    """Stable error categories retained from engine exceptions.

    These categories are intentionally broad so that engine-specific
    exceptions can be mapped without losing the diagnostic shape.
    """

    INVALID_PLAN = "INVALID_PLAN"
    UNSUPPORTED_OPERATION = "UNSUPPORTED_OPERATION"
    VALIDATION_FAILED = "VALIDATION_FAILED"
    TRANSIENT = "TRANSIENT"
    EXECUTION_FAILURE = "EXECUTION_FAILURE"
    OUTPUT_PUBLICATION_FAILURE = "OUTPUT_PUBLICATION_FAILURE"
    RECOVERY_INCONSISTENCY = "RECOVERY_INCONSISTENCY"
    CANCELLATION = "CANCELLATION"
    EXECUTION_LIMIT_EXCEEDED = "EXECUTION_LIMIT_EXCEEDED"
    UNKNOWN = "UNKNOWN"


class WorkflowError(Exception):
    """Unified error contract for the workflow orchestrator.

    Keeps the information a caller needs for retry decisions:
    category (for retry policy), step (for attribution), and diagnostics
    (engine-specific detail that must not be silently dropped).
    """

    def __init__(
        self,
        message: str,
        category: WorkflowErrorCategory,
        step: Optional[str] = None,
        retryable: bool = False,
        diagnostics: Optional[dict[str, Any]] = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.category = category
        self.step = step
        self.retryable = retryable
        self.diagnostics = diagnostics or {}


class StepResult:
    """Result contract for one executed step."""

    def __init__(
        self,
        status: StepStatus,
        output: Any = None,
        artifact: Optional[str] = None,
        retry_attempt: int = 0,
        max_retries: int = 1,
        diagnostics: Optional[dict[str, Any]] = None,
    ) -> None:
        self.status = status
        self.output = output
        self.artifact = artifact
        self.retry_attempt = retry_attempt
        self.max_retries = max_retries
        self.diagnostics = diagnostics or {}

    @staticmethod
    def _json_safe_value(obj):
        """Coerce an arbitrary object into a JSON-safe form so persistence of
        step results never crashes (SpreadsheetAnalysis, file objects, etc.).
        Objects with to_dict() are expanded; everything else is carried as a
        small diagnostic payload so no data is silently dropped."""
        if obj is None or isinstance(obj, (str, int, float, bool)):
            return obj
        if isinstance(obj, (list, tuple)):
            return [StepResult._json_safe_value(v) for v in obj]
        if isinstance(obj, dict):
            return {str(k): StepResult._json_safe_value(v) for k, v in obj.items()}
        to_dict = getattr(obj, "to_dict", None)
        if callable(to_dict):
            return StepResult._json_safe_value(to_dict())
        return {
            "__class__": type(obj).__name__,
            "__module__": type(obj).__module__,
            "__repr__": repr(obj),
        }

    def to_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "output": StepResult._json_safe_value(self.output),
            "artifact": StepResult._json_safe_value(self.artifact),
            "retry_attempt": self.retry_attempt,
            "max_retries": self.max_retries,
            "diagnostics": StepResult._json_safe_value(self.diagnostics),
        }


@dataclass
class ProvenanceRecord:
    """Per-step provenance so recovery and audit can be traced."""

    executed_at: str
    step: str
    engine: str
    status: str
    input_path: Optional[str] = None
    output_path: Optional[str] = None
    retry_attempt: int = 0
    category: Optional[str] = None
    message: str = ""


@dataclass
class WorkflowContext:
    """Runtime context a workflow runner carries through execution."""

    workflow_id: str
    plan: "WorkflowPlan"
    status: WorkflowStatus = WorkflowStatus.PLANNED
    input: Any = None
    policy: dict[str, Any] = field(default_factory=dict)
    limits: dict[str, Any] = field(default_factory=dict)
    created_at: str = field(default_factory=_now_iso)
    updated_at: str = field(default_factory=_now_iso)

    def to_dict(self) -> dict[str, Any]:
        return {
            "workflow_id": self.workflow_id,
            "status": self.status.value,
            "plan": self.plan.to_dict(),
            "input": self.input,
            "policy": self.policy,
            "limits": self.limits,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }


@dataclass
class WorkflowPlan:
    """Deterministic plan (rule-based by default; AI optional on top)."""

    workflow_id: str
    schema_version: str = "1.0.0"
    created_at: str = field(default_factory=_now_iso)
    steps: list["Step"] = field(default_factory=list)
    output_policy: dict[str, Any] = field(default_factory=dict)
    limits: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "workflow_id": self.workflow_id,
            "schema_version": self.schema_version,
            "created_at": self.created_at,
            "steps": [step.to_dict() for step in self.steps],
            "output_policy": self.output_policy,
            "limits": self.limits,
        }


@dataclass
class Step:
    """One atomic operation plus its local execution state."""

    id: str
    operation: str
    parameters: dict[str, Any] = field(default_factory=dict)
    dependencies: list[str] = field(default_factory=list)
    status: StepStatus = StepStatus.PENDING
    result: Optional[StepResult] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "operation": self.operation,
            "parameters": self.parameters,
            "dependencies": self.dependencies,
            "status": self.status.value,
            "result": self.result.to_dict() if self.result else None,
        }
