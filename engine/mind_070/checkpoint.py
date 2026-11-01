"""MIND-070C checkpointing and recovery.

Safe boundaries:
- A checkpoint is written ONLY after a step has fully succeeded
  (post-execution, pre-next-step). A failed step never produces a checkpoint.
- All checkpoint writes use the crash-safe atomic writer from engine.mind_070.state.

Recovery contract:
- Loading always validates schema version + required fields. Corrupted or
  incompatible state becomes a WorkflowError (never silently accepted).
- The classification determines what recovery may do:
    RECOVERED              state + full checkpoint reached, resumable/complete
    PARTIAL                some steps completed, more can be resumed
    NEEDS_RECOVERY         state present but incomplete/uncertain (no auto-repeat)
    FAILED / NO_STATE      nothing usable to resume from
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any, Optional

from engine.mind_070.contracts import WorkflowError, WorkflowErrorCategory, WorkflowStatus
from engine.mind_070.state import (
    _CHECKPOINT_FILE,
    _STATE_FILE,
    _now_iso,
    _CHECKPOINT_FILE,
    _STATE_FILE,
    _now_iso,
    CheckpointRecord,
    SCHEMA_VERSION,
    WorkflowState,
    _write_json_atomic,
    checkpoint_file_path,
    delete_workflow_state,
    load_workflow_state,
    save_workflow_state,
    state_file_path,
)

__all__ = [
    "CheckpointRecord",
    "RecoveryResult",
    "CheckpointStatus",
    "recovery_status",
    "classify_recovery",
    "write_checkpoint",
    "write_workflow_state",
    "load_checkpoint",
    "checkpoint_exists",
    "delete_checkpoint",
    "recover_workflow",
]


@dataclass(frozen=True)
class RecoveryResult:
    workflow_id: str
    classification: str
    state: Optional["WorkflowState"] = None
    error: Optional[str] = None
    last_checkpoint: Optional["CheckpointRecord"] = None


@dataclass(frozen=True)
class CheckpointStatus:
    SUCCESS = "SUCCESS"
    PARTIAL = "PARTIAL"
    FAILED = "FAILED"
    NEEDS_RECOVERY = "NEEDS_RECOVERY"
    UNKNOWN = "UNKNOWN"


def recovery_status(ws: Optional[WorkflowState]) -> str:
    if ws is None:
        return CheckpointStatus.UNKNOWN
    if ws.status == WorkflowStatus.SUCCEEDED:
        return CheckpointStatus.SUCCESS
    if ws.status == WorkflowStatus.PARTIAL:
        return CheckpointStatus.PARTIAL
    if ws.status == WorkflowStatus.FAILED:
        return CheckpointStatus.FAILED
    if ws.status == WorkflowStatus.CANCELLED:
        return CheckpointStatus.PARTIAL
    return CheckpointStatus.NEEDS_RECOVERY


def classify_recovery(ws: Optional[WorkflowState]) -> str:
    if ws is None:
        return "NO_STATE"
    if ws.status == WorkflowStatus.SUCCEEDED and _missing_artifact_detected(ws):
        return "NEEDS_RECOVERY"
    if ws.status == WorkflowStatus.SUCCEEDED:
        return "RECOVERED"
    if ws.status == WorkflowStatus.PARTIAL:
        return "PARTIAL"
    if ws.status == WorkflowStatus.FAILED:
        return "FAILED"
    if ws.status == WorkflowStatus.CANCELLED:
        return "CANCELLED"
    return "NEEDS_RECOVERY"


def _missing_artifact_detected(ws: Optional[WorkflowState]) -> bool:
    """Whether any completed artifact or referenced path is missing."""
    if ws is None:
        return False
    for step in ws.steps:
        if step.status == WorkflowStatus.SUCCEEDED and step.artifact:
            if not os.path.exists(step.artifact):
                return True
    for checkpoint in ws.checkpoints:
        for ref in (checkpoint.artifact_refs or []):
            if ref and not os.path.exists(ref):
                return True
    return False


def _latest_checkpoint(snapshot: Optional[WorkflowState]) -> Optional["CheckpointRecord"]:
    if not snapshot or not snapshot.checkpoints:
        return None
    records = sorted(snapshot.checkpoints, key=lambda c: c.saved_at)
    return records[-1]


def write_checkpoint(workflow_state: WorkflowState, workflow_id: str, base_dir: str = ".memind") -> str:
    """Write state.json + checkpoint.json at a safe (post-success) boundary.

    Callers must only invoke this after a step has fully succeeded.
    """
    last_step_id = ""
    last_artifact_refs: list[str] = []
    if workflow_state.provenance and workflow_state.provenance.records:
        last_provenance = workflow_state.provenance.records[-1]
        last_step_id = last_provenance.step_id
        last_artifact_refs = list(last_provenance.artifact_refs)
    if not last_step_id and workflow_state.steps:
        last_artifact_refs = [workflow_state.steps[-1].artifact or ""] if workflow_state.steps else []

    ckpt = CheckpointRecord(
        checkpoint_id=f"{workflow_id}@{_now_iso()}",
        step_id=last_step_id,
        status=workflow_state.status.value,
        summary={"completed_steps": len(workflow_state.steps)},
        artifact_refs=last_artifact_refs,
    )
    state_path = save_workflow_state(workflow_state, workflow_id, base_dir)
    ckpt_path = checkpoint_file_path(workflow_id, base_dir)
    _write_json_atomic(ckpt_path, ckpt.to_dict())
    # update the state's current checkpoint reference
    workflow_state.current_checkpoint = ckpt.checkpoint_id
    workflow_state.updated_at = _now_iso()
    save_workflow_state(workflow_state, workflow_id, base_dir)
    return ckpt_path


def write_workflow_state(workflow_state: WorkflowState, workflow_id: str, base_dir: str = ".memind") -> str:
    return save_workflow_state(workflow_state, workflow_id, base_dir)


def load_checkpoint(workflow_id: str, base_dir: str = ".memind") -> Optional[CheckpointRecord]:
    path = checkpoint_file_path(workflow_id, base_dir)
    if not os.path.exists(path):
        return None
    try:
        payload = _json_read(path)
    except (OSError, ValueError):
        return None
    if not isinstance(payload, dict) or payload.get("schema_version") != "1.0.0":
        return None
    try:
        return CheckpointRecord.from_dict(payload)
    except WorkflowError:
        return None


def checkpoint_exists(workflow_id: str, base_dir: str = ".memind") -> bool:
    return os.path.exists(checkpoint_file_path(workflow_id, base_dir))


def delete_checkpoint(workflow_id: str, base_dir: str = ".memind") -> None:
    try:
        os.unlink(checkpoint_file_path(workflow_id, base_dir))
    except FileNotFoundError:
        pass
    except OSError as exc:
        raise WorkflowError(f"cannot delete checkpoint {checkpoint_file_path(workflow_id, base_dir)}: {exc}", category=WorkflowErrorCategory.RECOVERY_FAILURE, step="<checkpoint>", retryable=False) from exc


def _json_read(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def recover_workflow(
    workflow_id: str,
    base_dir: str = ".memind",
    resume_step_index: Optional[int] = None,
) -> "RecoveryResult":
    """Load, validate and classify a persisted workflow state (Phase C).

    Returns a RecoveryResult. Corrupted/incompatible state is never accepted:
    it is reported as classification INVALID_SCHEMA_OR_STATE with the error.
    """
    from engine.mind_070.state import load_workflow_state

    path = state_file_path(workflow_id, base_dir)
    if not os.path.exists(path):
        return RecoveryResult(workflow_id=workflow_id, classification="NO_STATE")

    try:
        snapshot = load_workflow_state(workflow_id, base_dir)
    except WorkflowError as exc:
        return RecoveryResult(
            workflow_id=workflow_id,
            classification="INVALID_SCHEMA_OR_STATE",
            error=str(exc),
            last_checkpoint=_latest_checkpoint(None),  # type: ignore[arg-type]
        )
    except Exception as exc:  # safety net
        return RecoveryResult(
            workflow_id=workflow_id,
            classification="INVALID_STATE",
            error=str(exc),
            last_checkpoint=_latest_checkpoint(None),
        )

    classification = classify_recovery(snapshot)
    last_checkpoint = _latest_checkpoint(snapshot)

    result = RecoveryResult(
        workflow_id=workflow_id,
        classification=classification,
        state=snapshot,
        error=None,
        last_checkpoint=last_checkpoint,
    )
    if result.state is not None and resume_step_index is not None:
        # Reconstruct the recoverable frontier without silently re-executing
        # steps whose effect is uncertain.
        result = RecoveryResult(
            workflow_id=result.workflow_id,
            classification="NEEDS_RECOVERY" if result.classification in ("RECOVERED", "PARTIAL") and resume_step_index < len(result.state.steps) else result.classification,
            state=result.state,
            error=result.error,
            last_checkpoint=result.last_checkpoint,
        )
    return result
