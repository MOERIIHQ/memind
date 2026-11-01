"""Package entry point for MIND-070 (Specification v1.0).

Phase A (contracts, unified errors, registry, validator) and Phase B
(executor, adapters) live here. Phases C/D/E are separate, pending approval.
"""

from __future__ import annotations

# --- Phase A contracts & errors ---
from engine.mind_070.adapter import EngineAdapter
from engine.mind_070.contracts import (
    ProvenanceRecord,
    Step,
    StepResult,
    WorkflowContext,
    WorkflowError,
    WorkflowErrorCategory,
    WorkflowPlan,
    WorkflowStatus,
    StepStatus,
)
from engine.mind_070.operation_registry import (
    OperationInfo,
    OperationRegistry,
    OperationType,
    get_default_registry,
)
from engine.mind_070.plan_validator import (
    PlanValidationResult,
    ValidationError,
    validate_plan,
)

# --- Phase B executor & adapters ---
from engine.mind_070.adapters import (
    adapt_auto_reporting,
    adapt_file_enhance,
    adapt_formula_apply,
    adapt_spreadsheet_intelligence,
    adapt_text_task_to_formula,
)
from engine.mind_070.executor import WorkflowResult, execute_step, run_workflow
from engine.mind_070.cancellation import CancellationToken, cancel_requested
from engine.mind_070.limits import (
    GlobalLimits, WorkflowLimits, validate_limits_config, merge_limits,
)
from engine.mind_070.publisher import (
    OutputPublication, PublicationError, OVERWRITE_NEVER,
    OVERWRITE_ALWAYS, OVERWRITE_EXPLICIT,
)
from engine.mind_070.operation_adapter_registry import build_phase_b_registry

__all__ = [
    # Phase A
    "EngineAdapter",
    "ProvenanceRecord",
    "Step",
    "StepResult",
    "WorkflowContext",
    "WorkflowError",
    "WorkflowErrorCategory",
    "WorkflowPlan",
    "WorkflowStatus",
    "StepStatus",
    "OperationInfo",
    "OperationRegistry",
    "OperationType",
    "get_default_registry",
    "PlanValidationResult",
    "ValidationError",
    "validate_plan",
    # Phase B
    "WorkflowResult",
    "execute_step",
    "run_workflow",
    "adapt_text_task_to_formula",
    "adapt_formula_apply",
    "adapt_file_enhance",
    "adapt_spreadsheet_intelligence",
    "adapt_auto_reporting",
    "build_phase_b_registry",
    # Phase D
    "GlobalLimits", "WorkflowLimits", "validate_limits_config", "merge_limits",
    "CancellationToken", "cancel_requested",
    "OutputPublication", "PublicationError",
    "OVERWRITE_NEVER", "OVERWRITE_ALWAYS", "OVERWRITE_EXPLICIT",
]
