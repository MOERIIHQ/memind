"""MIND-070D execution limits: global + per-workflow.

Providers:
- GlobalLimits: defaults shared by every workflow in a deployment.
- WorkflowLimits: per-workflow overrides; a field that is None inherits the
  global value.

Retry contract (D3):
- ``GlobalLimits.max_retries`` is the configured cap on automatic retries.
  Default is 1 (one additional retry, at most two attempts). It is consumed by
  the executor's retry loop; it is not dead configuration.
- Workflow retry and AI retry are separate budgets (D3).

Validation rule (D.3): a mis-configured limit must raise and must NOT silently
disarm the limit enforcement. If the shell cannot be parsed, the system raises
WorkflowError and leaves enforcement at the previously loaded config.

Sizes are provided in code as provisional Phase-1 values (D4). They become
final only after a benchmark on the target device and explicit approval.
"""


from __future__ import annotations

import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Optional

from engine.mind_070.contracts import WorkflowError, WorkflowErrorCategory

__all__ = [
    "GlobalLimits", "WorkflowLimits", "DEFAULT_GLOBAL_LIMITS",
    "validate_limits_config", "merge_limits", "check_workflow_limits",
    "StepLimitRecord", "CheckStatus",
    # Provisional candidate bounds (D4). Do not treat as final.
    "MAX_CANDIDATE_STEPS", "MAX_CANDIDATE_JSON_BYTES",
    "MAX_PROMPT_BYTES", "MAX_AI_RESPONSE_BYTES",
    # Provisional execution bounds (D4).
    "MAX_INPUT_BYTES", "MAX_OUTPUT_BYTES",
]

# ---------------------------------------------------------------------------
# dataclasses
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class GlobalLimits:
    max_steps: int
    max_duration_seconds: float
    max_step_duration_seconds: float
    max_input_size_bytes: int
    max_output_size_bytes: int
    max_retries: int

    @classmethod
    def defaults(cls) -> "GlobalLimits":
        # Phase-1 provisional values (D4).
        # Input/output are no longer "unlimited"; they are now safety bounds.
        max_input = 32 * 1024 * 1024   # 32 MiB
        max_output = 64 * 1024 * 1024  # 64 MiB
        # One retry adds a second attempt, at most two attempts total.
        return cls(
            max_steps=120,
            max_duration_seconds=3600.0,
            max_step_duration_seconds=300.0,
            max_input_size_bytes=max_input,
            max_output_size_bytes=max_output,
            max_retries=1,
        )


@dataclass(frozen=True)
class WorkflowLimits:
    max_steps: Optional[int] = None
    max_duration_seconds: Optional[float] = None
    max_step_duration_seconds: Optional[float] = None
    max_input_size_bytes: Optional[int] = None
    max_output_size_bytes: Optional[int] = None
    max_retries: Optional[int] = None
    max_ai_response_bytes: Optional[int] = None
    reason: str = ""

    def apply_overrides(self, global_limits: GlobalLimits) -> dict[str, Any]:
        resolved: dict[str, Any] = {
            "max_steps": global_limits.max_steps,
            "max_duration_seconds": global_limits.max_duration_seconds,
            "max_step_duration_seconds": global_limits.max_step_duration_seconds,
            "max_input_size_bytes": global_limits.max_input_size_bytes,
            "max_output_size_bytes": global_limits.max_output_size_bytes,
            "max_retries": global_limits.max_retries,
            "max_ai_response_bytes": 0,
        }
        for key in resolved:
            if getattr(self, key) is not None:
                resolved[key] = getattr(self, key)
        return resolved


DEFAULT_GLOBAL_LIMITS = GlobalLimits.defaults()


# ---------------------------------------------------------------------------
# Provisional Phase-1 candidate bounds (D4)
# ---------------------------------------------------------------------------
MAX_CANDIDATE_STEPS = 120             # kandidat plan: maks 120 langkah
MAX_CANDIDATE_JSON_BYTES = 1024 * 1024  # kandidat plan JSON: maks 1 MiB
MAX_PROMPT_BYTES = 32 * 1024              # AI prompt: maks 32 KiB
MAX_AI_RESPONSE_BYTES = 1024 * 1024       # AI response: maks 1 MiB

# Provisional execution bounds (D4).
MAX_INPUT_BYTES = 32 * 1024 * 1024    # input file: 32 MiB
MAX_OUTPUT_BYTES = 64 * 1024 * 1024   # output file: 64 MiB


# ---------------------------------------------------------------------------
# validation
# ---------------------------------------------------------------------------

def validate_limits_config(config: Any) -> GlobalLimits:
    """Parse a limits config into GlobalLimits; raise on any invalid value.

    A bad config raises WorkflowError and does NOT disable enforcement. The
    caller is expected to record the reason and stop the workflow.
    """
    if config is None:
        return DEFAULT_GLOBAL_LIMITS
    if isinstance(config, dict):
        raw = config
    elif isinstance(config, (list, tuple)):
        raw = {"max_steps": config}  # not used; keep strict below
        raise WorkflowError(
            "limits config must be a dict", category=WorkflowErrorCategory.INVALID_PLAN
        )
    else:
        raise WorkflowError(
            "limits config must be a dict", category=WorkflowErrorCategory.INVALID_PLAN
        )

    cls = GlobalLimits
    fields = {
        "max_steps": ("max_steps", int),
        "max_duration_seconds": ("max_duration_seconds", (int, float)),
        "max_step_duration_seconds": ("max_step_duration_seconds", (int, float)),
        "max_input_size_bytes": ("max_input_size_bytes", int),
        "max_output_size_bytes": ("max_output_size_bytes", int),
        "max_retries": ("max_retries", int),
    }
    resolved: dict[str, Any] = {}
    for attr, (key, typ) in fields.items():
        value = raw.get(key, cls.defaults().__getattribute__(attr))
        if isinstance(value, bool) or not isinstance(value, typ):
            raise WorkflowError(
                f"limits.{key} must be {typ.__name__}, got {type(value).__name__}",
                category=WorkflowErrorCategory.INVALID_PLAN,
            )
        if value < 0:
            raise WorkflowError(
                f"limits.{key} must be >= 0", category=WorkflowErrorCategory.INVALID_PLAN
            )
        resolved[attr] = value
    return cls(**resolved)


def merge_limits(workflow_limits: Any, global_limits: GlobalLimits) -> dict[str, Any]:
    """Resolve per-workflow overrides over global limits (global wins when
    per-workflow is None)."""
    if workflow_limits is None or isinstance(workflow_limits, (int, float, bool)):
        return global_limits.__dict__
    if isinstance(workflow_limits, dict):
        data = workflow_limits
    elif isinstance(workflow_limits, WorkflowLimits):
        data = workflow_limits
    else:
        raise WorkflowError(
            "per-workflow limits must be a dict or WorkflowLimits",
            category=WorkflowErrorCategory.INVALID_PLAN,
        )
    sections: dict[str, Any] = {}
    for attr in GlobalLimits.__dataclass_fields__:
        global_value = getattr(global_limits, attr)
        if isinstance(data, dict):
            workflow_value = data.get(attr, None)
        else:
            workflow_value = getattr(data, attr, None)
        sections[attr] = global_value if workflow_value is None else workflow_value
    return sections


# ---------------------------------------------------------------------------
# enforcement and measurement
# ---------------------------------------------------------------------------

class StepLimitRecord:
    def __init__(self, step_id: str, start: float, input_size: int, output_size: int) -> None:
        self.step_id = step_id
        self.start = start
        self.input_size = input_size
        self.output_size = output_size
        self.elapsed = 0.0
        self.recorded = False


def check_workflow_limits(
    limits: dict[str, Any],
    step_index: int,
    start_time: float,
    input_size: int = 0,
    output_size: int = 0,
    max_ai_response_bytes: int = 0,
) -> None:
    """Raise WorkflowError(EXECUTION_LIMIT_EXCEEDED) when a limit is breached.

    Failure of configuration is reported here (not silently ignored); the
    caller records the final status and exits cleanly.
    """
    if limits.get("max_steps", 0) and step_index > limits["max_steps"]:
        raise WorkflowError(
            f"global limit exceeded: step count {step_index} > max_steps {limits['max_steps']}",
            category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="<workflow>",
            retryable=False,
            diagnostics={"max_steps": limits["max_steps"], "step_index": step_index},
        )
    if limits.get("max_duration_seconds", 0) and (time.monotonic() - start_time) > limits["max_duration_seconds"]:
        raise WorkflowError(
            f"global limit exceeded: workflow duration {time.monotonic() - start_time:.3f}s > {limits['max_duration_seconds']}s",
            category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="<workflow>",
            retryable=False,
            diagnostics={"max_duration_seconds": limits["max_duration_seconds"]},
        )
    if limits.get("max_input_size_bytes", 0) and input_size > limits["max_input_size_bytes"]:
        raise WorkflowError(
            f"input size {input_size} > max_input_size_bytes {limits['max_input_size_bytes']}",
            category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="<workflow>",
            retryable=False,
            diagnostics={"max_input_size_bytes": limits["max_input_size_bytes"], "input_size": input_size},
        )
    if limits.get("max_output_size_bytes", 0) and output_size > limits["max_output_size_bytes"]:
        raise WorkflowError(
            f"output size {output_size} > max_output_size_bytes {limits['max_output_size_bytes']}",
            category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="<workflow>",
            retryable=False,
            diagnostics={"max_output_size_bytes": limits["max_output_size_bytes"], "output_size": output_size},
        )
    if max_ai_response_bytes > 0:
        # AI output size; measured by the AIProvider layer (future phase).
        pass
