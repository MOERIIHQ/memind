"""MIND-070C workflow state persistence and recovery (JSON).

Site:
  <base_dir>/<workflow_id>/state.json      (workflow state, crash-safe)
  <base_dir>/<workflow_id>/checkpoint.json (checkpoint payload, crash-safe)

Rules applied here:
- Every public function that writes a file goes through _write_json_atomic()
  (temp file in the same directory + os.replace), so a failure never leaves a
  half-written state file.
- Reading always runs schema + required-field validation. A corrupted,
  non-object, or schema-incompatible file becomes a WorkflowError instead of
  being silently accepted.
- No runtime code in this module reads live workbook contents; provenance only
  stores IDs/diagnostics, never credentials nor workbook payloads.
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Optional

from engine.mind_070.contracts import (
    Step,
    StepResult,
    StepStatus,
    WorkflowError,
    WorkflowErrorCategory,
    WorkflowPlan,
    WorkflowStatus,
)
from engine.mind_070.provenance import ProvenanceRecord, StepHistory

SCHEMA_VERSION = "1.0.0"

_REQUIRED_KEYS = frozenset(
    {
        "schema_version",
        "workflow_id",
        "status",
        "created_at",
        "updated_at",
        "steps",
        "provenance",
        "checkpoints",
        "plan_version",
    }
)

_STATE_FILE = "state.json"
_CHECKPOINT_FILE = "checkpoint.json"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _to_step_status(value: Any) -> Optional[StepStatus]:
    if value is None:
        return None
    if isinstance(value, StepStatus):
        return value
    return StepStatus(str(value))


def _to_workflow_status(value: Any) -> Optional[WorkflowStatus]:
    if value is None:
        return None
    if isinstance(value, WorkflowStatus):
        return value
    return WorkflowStatus(str(value))


def _out_path(base_dir: str, workflow_id: str, name: str) -> str:
    os.makedirs(os.path.join(base_dir, workflow_id), exist_ok=True)
    return os.path.join(base_dir, workflow_id, name)


def _write_json_atomic(path: str, payload: Any) -> None:
    directory = os.path.dirname(path)
    os.makedirs(directory, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=directory, prefix=".tmp-", suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _read_json(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def schema_version(payload: Any) -> str:
    if not isinstance(payload, dict):
        raise WorkflowError("state file is not a JSON object", category=WorkflowErrorCategory.INVALID_PLAN, step="<state>", retryable=False)
    stored = payload.get("schema_version")
    if stored != SCHEMA_VERSION:
        raise WorkflowError(
            f"incompatible schema version: {stored!r} != {SCHEMA_VERSION!r}",
            category=WorkflowErrorCategory.RECOVERY_INCONSISTENCY,
            step="<state>",
            retryable=False,
        )
    return stored


def _validate_payload(payload: Any) -> None:
    if not isinstance(payload, dict):
        raise WorkflowError("state file is not a JSON object", category=WorkflowErrorCategory.INVALID_PLAN, step="<state>", retryable=False)
    missing = _REQUIRED_KEYS - set(payload.keys())
    if missing:
        raise WorkflowError(f"missing required state field(s): {sorted(missing)}", category=WorkflowErrorCategory.INVALID_PLAN, step="<state>", retryable=False)
    schema_version(payload)


@dataclass
class CheckpointRecord:
    checkpoint_id: str
    step_id: str
    status: str
    summary: Optional[dict] = None
    artifact_refs: list = field(default_factory=list)
    saved_at: str = field(default_factory=_now_iso)

    def to_dict(self) -> dict:
        return {
            "schema_version": "1.0.0",
            "checkpoint_id": str(self.checkpoint_id),
            "step_id": str(self.step_id),
            "status": str(self.status),
            "summary": self.summary,
            "artifact_refs": list(self.artifact_refs),
            "saved_at": str(self.saved_at),
        }

    @classmethod
    def from_dict(cls, data: Any) -> "CheckpointRecord":
        if not isinstance(data, dict):
            raise WorkflowError("checkpoint record is not a JSON object", category=WorkflowErrorCategory.INVALID_PLAN, step="<checkpoint>", retryable=False)
        missing = {"checkpoint_id", "step_id", "status"} - set(data.keys())
        if missing:
            raise WorkflowError(f"checkpoint record missing required field(s): {sorted(missing)}", category=WorkflowErrorCategory.INVALID_PLAN, step="<checkpoint>", retryable=False)
        return cls(
            checkpoint_id=str(data["checkpoint_id"]),
            step_id=str(data["step_id"]),
            status=str(data["status"]),
            summary=data.get("summary"),
            artifact_refs=list(data.get("artifact_refs", []) or []),
            saved_at=str(data.get("saved_at", _now_iso())),
        )


@dataclass
class WorkflowState:
    workflow_id: str
    input: Any = None
    plan: WorkflowPlan = field(default_factory=lambda: WorkflowPlan(workflow_id=""))
    status: WorkflowStatus = WorkflowStatus.PLANNED
    steps: list[StepResult] = field(default_factory=list)
    provenance: "StepHistory" = field(default_factory=StepHistory)
    checkpoints: list[CheckpointRecord] = field(default_factory=list)
    policy: dict = field(default_factory=dict)
    limits: dict = field(default_factory=dict)
    current_checkpoint: Optional[str] = None
    created_at: str = field(default_factory=_now_iso)
    updated_at: str = field(default_factory=_now_iso)

    def to_dict(self) -> dict:
        return {
            "schema_version": SCHEMA_VERSION,
            "workflow_id": self.workflow_id,
            "status": self.status.value,
            "input": self.input,
            "plan": self.plan.to_dict() if self.plan else None,
            "steps": [step.to_dict() for step in self.steps],
            "provenance": self.provenance.to_dict() if self.provenance else None,
            "checkpoints": [checkpoint.to_dict() for checkpoint in self.checkpoints],
            "policy": dict(self.policy),
            "limits": dict(self.limits),
            "current_checkpoint": self.current_checkpoint,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "plan_version": self.plan.schema_version if self.plan else SCHEMA_VERSION,
        }

    @classmethod
    def from_dict(cls, data: Any) -> "WorkflowState":
        if not isinstance(data, dict):
            raise WorkflowError("workflow state is not a JSON object", category=WorkflowErrorCategory.INVALID_PLAN, step="<state>", retryable=False)
        _validate_payload(data)
        plan = data.get("plan")
        plan_obj = None
        if plan:
            plan_obj = _restore_plan(plan)
        provenance_data = data.get("provenance")
        provenance = StepHistory.from_dict(provenance_data) if provenance_data else StepHistory()
        checkpoints = [CheckpointRecord.from_dict(c) for c in (data.get("checkpoints") or [])]
        steps = [_restore_step_result(s) for s in (data.get("steps") or [])]
        return cls(
            workflow_id=str(data["workflow_id"]),
            input=data.get("input"),
            plan=plan_obj,
            status=_to_workflow_status(data["status"]),
            steps=steps,
            provenance=provenance,
            checkpoints=checkpoints,
            policy=dict(data.get("policy") or {}),
            limits=dict(data.get("limits") or {}),
            current_checkpoint=data.get("current_checkpoint"),
            created_at=str(data.get("created_at", _now_iso())),
            updated_at=str(data.get("updated_at", _now_iso())),
        )


def _restore_step_result(raw: Any) -> Any:
    if isinstance(raw, StepResult):
        return raw
    if not isinstance(raw, dict):
        return raw
    return StepResult(
        status=_to_step_status(raw.get("status")),
        output=raw.get("output"),
        artifact=raw.get("artifact"),
        retry_attempt=int(raw.get("retry_attempt", 0)),
        max_retries=int(raw.get("max_retries", 1)),
        diagnostics=dict(raw.get("diagnostics") or {}),
    )


def _restore_step(raw: Any) -> Any:
    if isinstance(raw, Step):
        return raw
    if not isinstance(raw, dict):
        return raw
    return Step(
        id=_restore_step_id(raw.get("id")),
        operation=str(raw.get("operation")),
        parameters=raw.get("parameters"),
        dependencies=list(raw.get("dependencies") or []),
        status=_to_step_status(raw.get("status")),
        result=_restore_step_result(raw.get("result")),
    )


def _restore_step_id(raw: Any) -> str:
    return str(raw) if raw is not None else ""


def _restore_plan(raw: Any) -> Optional[WorkflowPlan]:
    if not raw:
        return None
    if not isinstance(raw, dict):
        return raw
    steps = [_restore_step(sd) for sd in (raw.get("steps") or [])]
    return WorkflowPlan(
        workflow_id=str(raw.get("workflow_id")),
        schema_version=str(raw.get("schema_version")),
        created_at=raw.get("created_at"),
        steps=steps,
        output_policy=dict(raw.get("output_policy") or {}),
        limits=dict(raw.get("limits") or {}),
    )


def state_file_path(workflow_id: str, base_dir: str = ".memind") -> str:
    return _out_path(base_dir, workflow_id, _STATE_FILE)


def checkpoint_file_path(workflow_id: str, base_dir: str = ".memind") -> str:
    return _out_path(base_dir, workflow_id, _CHECKPOINT_FILE)


def save_workflow_state(workflow_state: WorkflowState, workflow_id: str, base_dir: str = ".memind") -> str:
    path = state_file_path(workflow_id, base_dir)
    _write_json_atomic(path, workflow_state.to_dict())
    return path


def load_workflow_state(workflow_id: str, base_dir: str = ".memind") -> WorkflowState:
    path = state_file_path(workflow_id, base_dir)
    try:
        payload = _read_json(path)
    except (OSError, ValueError) as exc:
        raise WorkflowError(
            f"cannot read workflow state {path}: {exc}",
            category=WorkflowErrorCategory.RECOVERY_INCONSISTENCY,
            step="<state>",
            retryable=False,
        ) from exc
    _validate_payload(payload)
    return WorkflowState.from_dict(payload)


def delete_workflow_state(workflow_id: str, base_dir: str = ".memind") -> None:
    path = state_file_path(workflow_id, base_dir)
    try:
        os.unlink(path)
    except FileNotFoundError:
        pass
    except OSError as exc:
        raise WorkflowError(f"cannot delete workflow state {path}: {exc}", category=WorkflowErrorCategory.RECOVERY_FAILURE, step="<state>", retryable=False) from exc
