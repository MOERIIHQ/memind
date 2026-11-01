"""MIND-070C provenance and step history (JSON round-trip enabled).

Tracks, per step, the engine used, status, artifacts, retry attempt, executed
time and diagnostics. Provenance records are stored in the workflow state JSON
and are safe to round-trip: the dataclasses only hold plain, JSON-serializable
values plus the StepResult wrapper.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Optional

from engine.mind_070.contracts import StepResult, StepStatus, WorkflowError, WorkflowErrorCategory

__all__ = ["ProvenanceRecord", "StepHistory"]


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _to_step_status(value: Any) -> Optional[StepStatus]:
    if value is None:
        return None
    if isinstance(value, StepStatus):
        return value
    return StepStatus(str(value))


def _restore_step_result(payload: Any) -> Optional[StepResult]:
    from engine.mind_070.contracts import StepResult as _SR

    if payload is None:
        return None
    if isinstance(payload, _SR):
        return payload.to_dict()
    if not isinstance(payload, dict):
        return payload
    # JSON stores provenance.step_result as a plain dict -- return it unchanged
    # so the record survives a round-trip without reconstructing objects.
    return payload


@dataclass
class ProvenanceRecord:
    """Single step provenance entry."""

    step_id: str
    engine: str
    status: str
    result: Optional[StepResult] = None
    retry_attempt: int = 0
    artifact_refs: list = field(default_factory=list)
    executed_at: str = field(default_factory=_now_iso)
    diagnostics: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "schema_version": "1.0.0",
            "step_id": str(self.step_id),
            "engine": str(self.engine),
            "status": str(self.status),
            "result": _restore_step_result(self.result),
            "retry_attempt": int(self.retry_attempt),
            "artifact_refs": list(self.artifact_refs),
            "executed_at": str(self.executed_at),
            "diagnostics": dict(self.diagnostics),
        }

    @classmethod
    def from_dict(cls, data: Any) -> "ProvenanceRecord":
        if not isinstance(data, dict):
            raise WorkflowError("provenance record is not a JSON object", category=WorkflowErrorCategory.INVALID_PLAN, step="<provenance>", retryable=False)
        missing = {"step_id", "engine", "status"} - set(data.keys())
        if missing:
            raise WorkflowError(f"provenance record missing required field(s): {sorted(missing)}", category=WorkflowErrorCategory.INVALID_PLAN, step="<provenance>", retryable=False)
        return cls(
            step_id=str(data["step_id"]),
            engine=str(data["engine"]),
            status=str(data["status"]),
            result=_restore_step_result(data.get("result")),
            retry_attempt=int(data.get("retry_attempt", 0)),
            artifact_refs=list(data.get("artifact_refs", []) or []),
            executed_at=str(data.get("executed_at", _now_iso())),
            diagnostics=dict(data.get("diagnostics") or {}),
        )


@dataclass
class StepHistory:
    """Ordered provenance history for one workflow."""

    records: list[ProvenanceRecord] = field(default_factory=list)

    def append(self, record: ProvenanceRecord) -> None:
        self.records.append(record)

    def to_list(self) -> list[dict]:
        return [record.to_dict() for record in self.records]

    def to_dict(self) -> dict:
        return {"schema_version": "1.0.0", "records": self.to_list()}

    @classmethod
    def from_dict(cls, data: Any) -> "StepHistory":
        records: list[ProvenanceRecord] = []
        if isinstance(data, dict):
            for item in data.get("records", []) or []:
                records.append(ProvenanceRecord.from_dict(item))
        return cls(records=records)
