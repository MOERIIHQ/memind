"""MIND-070B minimal Workflow Executor and per-step adapters dispatcher.

Responsibilities (Phase B scope):
- Resolve a plan step's registered operation and bind the matching adapter.
- Execute steps in dependency order (topological sort).
- Classify adapter component errors into the MIND-070 unified error model.
- Return a WorkflowResult that never reports SUCCEEDED when a step failed.

Explicitly NOT implemented in Phase B (deferred to Phase C/D/E):
- State persistence / provenance / checkpointing.
- Full retry policy (only per-step bounds recorded, no automatic re-run).
- Global/workflow limits enforcement.
- Cancellation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

import os
import json
import time

from engine.mind_070.contracts import (
    Step,
    StepResult,
    StepStatus,
    WorkflowContext,
    WorkflowError,
    WorkflowErrorCategory,
    WorkflowStatus,
    WorkflowPlan,
    _now_iso,
)
from engine.mind_070.operation_registry import OperationInfo, OperationRegistry
from engine.mind_070.state import CheckpointRecord
from engine.mind_070.cancellation import CancellationToken, cancel_requested
from engine.mind_070.limits import (
    GlobalLimits, WorkflowLimits, merge_limits,
    StepLimitRecord, validate_limits_config,
    check_workflow_limits,
    MAX_CANDIDATE_STEPS, MAX_CANDIDATE_JSON_BYTES,
)
from engine.mind_070.plan_validator import validate_plan
from engine.mind_070.publisher import (
    OutputPublication, PublicationError, OVERWRITE_NEVER,
    OVERWRITE_ALWAYS, OVERWRITE_EXPLICIT,
)
from engine.mind_070.provenance import ProvenanceRecord, StepHistory

__all__ = [
    "execute_step",
    "run_workflow",
    "WorkflowResult",
    "recover_workflow",
    "write_checkpoint",
    "load_checkpoint",
    "classify_recovery",
]


# --------------------------------------------------------------- dataclass ---

@dataclass
class WorkflowResult:
    """Outcome of running a workflow. Not persisted in Phase B."""

    workflow_id: str
    plan: WorkflowPlan
    step_results: list = field(default_factory=list)
    executed_step_ids: list = field(default_factory=list)
    overall_status: WorkflowStatus = WorkflowStatus.PLANNED
    error: Optional["WorkflowError"] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "workflow_id": self.workflow_id,
            "status": self.overall_status.value,
            "step_results": [s.to_dict() for s in self.step_results],
            "error": self.error.message if self.error else None,
        }


# ------------------------------------------------------------ execution ---

def execute_step(context: WorkflowContext, step: Step, registry: Optional[OperationRegistry] = None) -> StepResult:
    """Execute a single plan step with its registered adapter.

    Returns a StepResult. Never raises for component errors (they are classified
    into the StepResult); raises for invalid plan contract (unsupported operation,
    missing adapter) to signal the runner to halt.
    """
    # 1. registered + enabled operation
    op_info = _operation_info(context, step.operation, registry)
    if op_info is None or not op_info.enabled:
        err = WorkflowError(
            f"unsupported operation: {step.operation}",
            category=WorkflowErrorCategory.UNSUPPORTED_OPERATION,
            step=step.id,
            retryable=False,
        )
        return StepResult(
            status=StepStatus.FAILED,
            output={"error": str(err), "code": err.category.value},
            artifact=None,
            retry_attempt=0,
            max_retries=1,
            diagnostics={"error": str(err), "category": err.category.value},
        )

    adapter = op_info.adapter
    if adapter is None:
        err = WorkflowError(
            f"no adapter bound for operation: {step.operation}",
            category=WorkflowErrorCategory.UNSUPPORTED_OPERATION,
            step=step.id,
            retryable=False,
        )
        raise err

    # 2. parameter contract (dict only)
    if not isinstance(step.parameters, dict):
        return StepResult(
            status=StepStatus.FAILED,
            output={"error": "step parameters must be a dict", "code": "VALIDATION_FAILED"},
            artifact=None,
            retry_attempt=0,
            max_retries=1,
            diagnostics={"error": "step parameters must be a dict", "category": "VALIDATION_FAILED"},
        )

    # 3. delegate
    try:
        return adapter(context, step)
    except WorkflowError as exc:
        return StepResult(
            status=StepStatus.FAILED,
            output={"error": str(exc), "code": exc.category.value},
            artifact=None,
            retry_attempt=0,
            max_retries=1,
            diagnostics={"error": str(exc), "category": exc.category.value},
        )
    except Exception as exc:  # pragma: no cover - defensive
        return StepResult(
            status=StepStatus.FAILED,
            output={"error": str(exc), "code": type(exc).__name__},
            artifact=None,
            retry_attempt=0,
            max_retries=1,
            diagnostics={"error": str(exc), "category": "EXECUTION_FAILURE"},
        )


def _resolve_params(params: Any, executed: dict[str, Any]) -> Any:
    """Resolve minimal ``{'$ref': <step_id>}`` placeholders to prior outputs.

    Overly lossy for JSON; serialization belongs to Phase C. Only direct
    ``$ref`` dicts are substituted so end-to-end chains (e.g. 050 -> 060)
    work without duplicating engine logic.
    """
    if isinstance(params, dict):
        out = {}
        for key, value in params.items():
            if isinstance(value, dict) and "$ref" in value and len(value) == 1:
                out[key] = executed.get(value["$ref"])
            else:
                out[key] = _resolve_params(value, executed)
        return out
    if isinstance(params, list):
        return [_resolve_params(item, executed) for item in params]
    return params


def _operation_info(context: WorkflowContext, name: str, registry) -> Optional[OperationInfo]:
    if registry is not None and hasattr(registry, "get_operation"):
        return registry.get_operation(name)
    # Fall back to the registry attached on the WorkflowContext (set by
    # run_workflow() and, in tests, manually). The Phase B executor carries the
    # live registry and never reads an empty policy dict.
    ctx_registry = getattr(context, "registry", None)
    if ctx_registry is not None and hasattr(ctx_registry, "get_operation"):
        return ctx_registry.get_operation(name)
    if isinstance(context, dict) and isinstance(context.get("registry"), OperationRegistry):
        return context["registry"].get_operation(name)
    return None


def _topological_order(steps: list[Step]) -> list[Step]:
    """Return steps in dependency order (dependencies first)."""
    ids = {s.id for s in steps}
    lookup = {s.id: s for s in steps}
    state = {s.id: 0 for s in steps}  # 0 unvisited, 1 visiting, 2 done
    order: list[Step] = []

    def visit(sid: str) -> None:
        if state[sid] == 2:
            return
        if state[sid] == 1:
            raise WorkflowError(f"circular dependency in plan (step '{sid}')", category=WorkflowErrorCategory.INVALID_PLAN, step=sid, retryable=False)
        state[sid] = 1
        for dep in lookup[sid].dependencies:
            if dep not in ids:
                raise WorkflowError(f"dependency not in plan: '{dep}'", category=WorkflowErrorCategory.INVALID_PLAN, step=sid, retryable=False)
            visit(dep)
        state[sid] = 2
        order.append(lookup[sid])

    for s in steps:
        if state[s.id] == 0:
            visit(s.id)
    return order


def _default_retry_policy() -> dict:
    """Phase D defaults: one retry, transient errors only."""
    return {"max_retries": 1, "retryable_categories": {"TRANSIENT"}, "mode": "transient_only"}


def _RetryPolicy(max_retries: int, retryable_categories: set[str], mode: str) -> dict:
    return {"max_retries": max_retries, "retryable_categories": retryable_categories, "mode": mode}


def _is_retryable(result: StepResult, policy: dict) -> bool:
    """Transient/retryable only. Permanent failures (validation, unsupported,
    publication, limits) are never re-executed automatically.

    StepResult carries no 'retryable' flag in Phase B; retry eligibility is
    derived from the step's executed category (diagnostics.category) and the
    retry-policy declaration.
    """
    if getattr(result, "retryable", False):
        return True
    category = (result.diagnostics or {}).get("category")
    if category in policy["retryable_categories"]:
        return True
    if policy.get("mode") == "always":
        return True
    return False


def _execute_with_retry(context, step, registry, policy) -> StepResult:
    """Run one step, re-executing only retryable failures (max policy limit)."""
    attempt = 0
    result = execute_step(context, step, registry)
    retry_category: Optional[str] = None
    while attempt < policy["max_retries"] and _is_retryable(result, policy):
        attempt += 1
        retry_category = result.diagnostics.get("category")
        result.retry_attempt = attempt
        result = execute_step(context, step, registry)
    if attempt > 0:
        result.retry_attempt = attempt
        if retry_category and "retry_category" not in result.diagnostics:
            result.diagnostics = dict(result.diagnostics)
            result.diagnostics["retry_category"] = retry_category
    return result


def _provenance_record(step: Step, result: StepResult) -> ProvenanceRecord:
    return ProvenanceRecord(
        step_id=step.id,
        engine=str(result.diagnostics.get("engine", "unknown")),
        status=result.status.value,
        retry_attempt=result.retry_attempt,
        artifact_refs=[result.artifact] if result.artifact else [],
        executed_at=_now_iso(),
        diagnostics=dict(result.diagnostics or {}),
    )


def _check_step_duration(elapsed: float, limits: dict) -> None:
    if limits.get("max_step_duration_seconds", 0) and elapsed > limits["max_step_duration_seconds"]:
        raise WorkflowError(
            f"step duration {elapsed:.3f}s exceeds limit {limits['max_step_duration_seconds']}s",
            category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="<step>",
            retryable=False,
            diagnostics={"max_step_duration_seconds": limits["max_step_duration_seconds"], "elapsed": elapsed},
        )


def _file_size(path: Optional[str]) -> int:
    if not path or not isinstance(path, str):
        return 0
    try:
        return os.path.getsize(path)
    except (OSError, TypeError):
        return 0


def _publish_outputs(steps, publication: OutputPublication, overwrite_policy: str, explicit_overwrite: bool) -> None:
    required = []
    for step in steps:
        params = step.parameters if isinstance(step.parameters, dict) else {}
        out = params.get("output_path")
        if out:
            required.append(out)
    if not required:
        return
    missing = publication.validate_required_outputs(required)
    if missing:
        raise PublicationError(f"publication failed: {missing}")
    for step in steps:
        params = step.parameters if isinstance(step.parameters, dict) else {}
        out = params.get("output_path")
        if not out:
            continue
        try:
            publication.publish(step.parameters.get("input_path", ""), out)
        except PublicationError as exc:
            raise PublicationError(f"publish failed for {out!r}: {exc}") from exc


def run_workflow(
    workflow_id: str,
    plan: WorkflowPlan,
    input: Any = None,
    policy: dict = None,
    registry: Optional[OperationRegistry] = None,
    checkpoint_base_dir: Optional[str] = None,
    *,
    cancellation_token: Optional[CancellationToken] = None,
    retry_policy: Optional[dict] = None,
    limits: Optional[WorkflowLimits] = None,
    output_publication: Optional[OutputPublication] = None,
    overwrite_policy: str = OVERWRITE_NEVER,
    explicit_overwrite: bool = False,
) -> WorkflowResult:
    """Execute a plan in dependency order.

    Phase D (all optional; defaults preserve Phase B behaviour):
    - limits: global defaults + optional per-workflow overrides.
    - retry: re-execute a step only when its failure is transient/retryable,
      up to max_retries.
    - output_publication: publish each step output before declaring
      SUCCEEDED. Overwrite is denied by default; allow explicit permission.
    - cancellation_token: cooperative, checked only at step boundaries.
    """
    if registry is None:
        from engine.mind_070.operation_adapter_registry import get_default_registry

        registry = get_default_registry()
    context = WorkflowContext(
        workflow_id=workflow_id,
        plan=plan,
        input=input,
        policy=policy or {},
        limits=limits,
    )
    context.registry = registry

    # ---- D2 central validation gate: never execute an invalid plan ------------
    if not plan.steps:
        # Preserve the previous empty-plan behaviour: no side effects and a
        # FAILED outcome reported at the entry point (previously returned by the
        # loop after scanning zero steps).
        return WorkflowResult(
            workflow_id=workflow_id,
            plan=plan,
            step_results=[],
            executed_step_ids=[],
            overall_status=WorkflowStatus.FAILED,
            error=WorkflowError(
                "empty plan", category=WorkflowErrorCategory.INVALID_PLAN,
                step="<plan>", retryable=False,
            ),
        )

    # Registry may be passed in by a caller, but structural validation happens
    # right after it is bound so that no plan can bypass the gate.
    if registry is None:
        from engine.mind_070.operation_adapter_registry import get_default_registry
        registry = get_default_registry()

    structural = validate_plan(registry, plan)
    if not structural.valid:
        raise WorkflowError(
            "invalid plan: " + "; ".join(e.message for e in structural.errors),
            category=WorkflowErrorCategory.INVALID_PLAN,
            step="<plan>",
            retryable=False,
        )

    # D4 provisional candidate-plan bounds. Enforced on every plan that reaches
    # the executor so an AI-produced candidate can never bypass validation.
    if len(plan.steps) > MAX_CANDIDATE_STEPS:
        raise WorkflowError(
            f"candidate plan exceeds max_steps {MAX_CANDIDATE_STEPS}: "
            f"{len(plan.steps)}",
            category=WorkflowErrorCategory.INVALID_PLAN,
            step="<plan>",
            retryable=False,
        )
    try:
        plan_json = json.dumps(plan.to_dict(), ensure_ascii=False, sort_keys=True)
    except (TypeError, ValueError) as exc:
        raise WorkflowError(
            f"candidate plan JSON is not serializable: {exc}",
            category=WorkflowErrorCategory.INVALID_PLAN,
            step="<plan>",
            retryable=False,
        ) from exc
    plan_bytes = len(plan_json.encode("utf-8"))
    if plan_bytes > MAX_CANDIDATE_JSON_BYTES:
        raise WorkflowError(
            f"candidate plan JSON exceeds {MAX_CANDIDATE_JSON_BYTES} bytes: "
            f"{plan_bytes}",
            category=WorkflowErrorCategory.INVALID_PLAN,
            step="<plan>",
            retryable=False,
        )

    ordered = _topological_order(plan.steps)
    if os.environ.get("MIND070D_DEBUG"):
        print(f"[debug] run_workflow plan.steps={len(plan.steps)} ordered={len(ordered)} ids={[x.id for x in ordered]}")
    step_results: list[StepResult] = []
    executed_steps: list[Step] = []
    executed_step_ids: list[str] = []
    provenance = StepHistory()
    executed: dict[str, Any] = {}

    global_limits = GlobalLimits.defaults()
    try:
        # No limits provided  -> permissive (Phase B behaviour unchanged).
        # A policy is provided -> resolve per-workflow overrides over global;
        # an invalid policy raises and the workflow is marked FAILED.
        resolved_limits = {} if limits is None else merge_limits(limits, global_limits)
    except WorkflowError as exc:
        # Configuration failure does not silently disable limits.
        return WorkflowResult(
            workflow_id=workflow_id,
            plan=plan,
            step_results=step_results,
            executed_step_ids=executed_step_ids,
            overall_status=WorkflowStatus.FAILED,
            error=WorkflowError(
                f"invalid limits configuration: {exc.message}",
                category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
                step="<limits>",
                retryable=False,
            ),
        )
    start_time = time.monotonic()
    if retry_policy is not None:
        retry = _RetryPolicy(**retry_policy)
    else:
        # D3: one additional retry (at most two attempts) by default. GlobalLimits
        # max_retries is the configured cap and is consumed here; the default cap
        # is 1 (one additional retry). A caller-supplied limits config may
        # override the cap explicitly.
        max_retries = GlobalLimits.defaults().max_retries
        if isinstance(resolved_limits, dict):
            configured = resolved_limits.get("max_retries")
            if configured is not None:
                max_retries = configured
        retry = _RetryPolicy(
            max_retries=max_retries,
            retryable_categories={"TRANSIENT"},
            mode="transient_only",
        )

    for idx, step in enumerate(ordered, start=1):
        if os.environ.get("MIND070D_DEBUG"):
            print(f"[debug] loop idx={idx} ordered_len={len(ordered)} id={step.id}")
        if os.environ.get("MIND070D_EVERY"):
            print(f"[debug] EVERY idx={idx} step={step.id} result={getattr(step,'id','?')} results_so_far={len(step_results)}")
        if cancel_requested(cancellation_token):
            return WorkflowResult(
                workflow_id=workflow_id,
                plan=plan,
                step_results=step_results,
                executed_step_ids=executed_step_ids,
                overall_status=WorkflowStatus.CANCELLED,
                error=WorkflowError(
                    "workflow cancelled at a safe boundary", category=WorkflowErrorCategory.CANCELLATION, step="<cancel>", retryable=False,
                ),
            )
        # Resolve ``$ref`` placeholders to prior outputs so chained steps (e.g.
        # 050 -> 060) bind to the live object the adapter expects. The resolved
        # parameters are bound to a fresh step; the original plan step is left
        # untouched so the persisted plan remains canonical.
        input_size = _file_size(step.parameters.get("input_path")) if isinstance(step.parameters, dict) else 0
        output_size = 0
        try:
            check_workflow_limits(resolved_limits, idx, start_time, input_size, output_size)
        except WorkflowError as exc:
            # a limit breach is a controlled halt (D.3): report FAILED with the
            # reason, never crash and never a false SUCCESS.
            return WorkflowResult(
                workflow_id=workflow_id, plan=plan, step_results=step_results,
                executed_step_ids=executed_step_ids, overall_status=WorkflowStatus.FAILED,
                error=WorkflowError(f"global limit exceeded: {exc.message}",
                                     category=exc.category, step="<limits>", retryable=False),
            )
        step_start = time.monotonic()
        resolved_params = _resolve_params(step.parameters, executed)
        step_to_run = Step(id=step.id, operation=step.operation, parameters=resolved_params, dependencies=step.dependencies)
        result = _execute_with_retry(context, step_to_run, registry, retry)
        if os.environ.get("MIND070D_DEBUG") and step.id == "report":
            print(f"[debug] report before resolve: params={step.parameters!r}")
            print(f"[debug] report executed['intel']={executed.get('intel')!r} type={type(executed.get('intel')).__name__}")
        if os.environ.get("MIND070D_DEBUG"):
            print(f"[debug] step {step.id} output_type={type(result.output).__name__}", end="")
            if hasattr(result.output, 'workbook') or hasattr(result.output, 'sheets'):
                print(" (SpreadsheetAnalysis)", end="")
            print()
        step_elapsed = time.monotonic() - step_start
        try:
            _check_step_duration(step_elapsed, resolved_limits)
        except WorkflowError as exc:
            return WorkflowResult(
                workflow_id=workflow_id, plan=plan, step_results=step_results,
                executed_step_ids=executed_step_ids, overall_status=WorkflowStatus.FAILED,
                error=WorkflowError(f"step duration limit exceeded: {exc.message}",
                                     category=exc.category, step="<step>", retryable=False),
            )
        provenance.append(_provenance_record(step, result))
        step_results.append(result)
        executed_steps.append(step)
        executed_step_ids.append(step.id)
        executed[step.id] = result.output
        if result.status == StepStatus.SUCCEEDED:
            _checkpoint_workflow_state(workflow_id, plan, step_results, executed_steps, checkpoint_base_dir)
        if result.status == StepStatus.FAILED:
            # no automatic re-execution beyond retry policy (uncertain effect);
            # stop and report FAILED (no false success).
            return WorkflowResult(
                workflow_id=workflow_id,
                plan=plan,
                step_results=step_results,
                executed_step_ids=executed_step_ids,
                overall_status=WorkflowStatus.FAILED,
                error=_extract_workflow_error(result),
            )
        if os.environ.get("MIND070D_DEBUG"):
            print(f"[debug]   loop body end idx={idx} step={step.id}")
        if result.status == StepStatus.SUCCEEDED:
            _checkpoint_workflow_state(workflow_id, plan, step_results, executed_steps, checkpoint_base_dir)
        if result.status == StepStatus.FAILED:
            return WorkflowResult(
                workflow_id=workflow_id,
                plan=plan,
                step_results=step_results,
                executed_step_ids=executed_step_ids,
                overall_status=WorkflowStatus.FAILED,
                error=WorkflowError(
                    f"step '{step.id}' exceeded step duration limit",
                    category=WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
                    step=step.id,
                    retryable=False,
                ),
            )

    if os.environ.get("MIND070D_DEBUG"):
        print(f"[debug] AFTER loop: len(step_results)={len(step_results)} len(executed_steps)={len(executed_steps)} ordered={len(ordered)}")
    # empty plan is never succeeded (no steps to produce the required output)
    if not step_results:
        return WorkflowResult(
            workflow_id=workflow_id,
            plan=plan,
            step_results=step_results,
            executed_step_ids=executed_step_ids,
            overall_status=WorkflowStatus.FAILED,
            error=WorkflowError(
                "empty plan", category=WorkflowErrorCategory.INVALID_PLAN, step="<plan>", retryable=False,
            ),
        )

    # ---- output publication (new) ----
    if output_publication is not None:
        try:
            _publish_outputs(plan.steps, output_publication, overwrite_policy, explicit_overwrite)
        except PublicationError as exc:
            provenance.append(
                ProvenanceRecord(
                    step_id="<publication>",
                    engine="MIND-070D-publisher",
                    status="FAILED",
                    retry_attempt=0,
                    artifact_refs=[],
                    executed_at=_now_iso(),
                    diagnostics={"error": str(exc), "category": WorkflowErrorCategory.OUTPUT_PUBLICATION_FAILURE.value},
                )
            )
            return WorkflowResult(
                workflow_id=workflow_id,
                plan=plan,
                step_results=step_results,
                executed_step_ids=executed_step_ids,
                overall_status=WorkflowStatus.FAILED,
                error=WorkflowError(
                    f"output publication failed: {exc}",
                    category=WorkflowErrorCategory.OUTPUT_PUBLICATION_FAILURE,
                    step="<publication>",
                    retryable=False,
                ),
            )

    if cancel_requested(cancellation_token):
        return WorkflowResult(
            workflow_id=workflow_id,
            plan=plan,
            step_results=step_results,
            executed_step_ids=executed_step_ids,
            overall_status=WorkflowStatus.CANCELLED,
            error=WorkflowError(
                "workflow cancelled at safe boundary after execution", category=WorkflowErrorCategory.CANCELLATION, step="<cancel>", retryable=False,
            ),
        )

    all_ok = all(r.status == StepStatus.SUCCEEDED for r in step_results)
    overall = WorkflowStatus.SUCCEEDED if all_ok else WorkflowStatus.PARTIAL
    return WorkflowResult(
        workflow_id=workflow_id,
        plan=plan,
        step_results=step_results,
        executed_step_ids=executed_step_ids,
        overall_status=overall,
        error=None,
    )


def _checkpoint_workflow_state(workflow_id: str, plan: WorkflowPlan, step_results: list[StepResult], executed_steps: list[Step], base_dir: Optional[str]) -> None:
    """Persist a crash-safe workflow state + checkpoint at a successful step.

    Called by run_workflow() whenever a step transitions to SUCCEEDED, so that
    the next step can resume from this checkpoint (Phase C). Provenance records
    use the executed Step (not the StepResult), which is the only object that
    carries the step identifier.
    """
    if not base_dir:
        return
    from engine.mind_070.checkpoint import write_checkpoint
    from engine.mind_070.provenance import ProvenanceRecord, StepHistory
    from engine.mind_070.state import WorkflowState

    records = []
    for step, result in zip(executed_steps, step_results):
        records.append(
            ProvenanceRecord(
                step_id=step.id,
                engine=str(result.diagnostics.get("engine", "unknown")),
                status=result.status.value,
                result=result,
                artifact_refs=[result.artifact] if result.artifact else [],
            )
        )
    checkpoints: list[CheckpointRecord] = [
        CheckpointRecord(
            checkpoint_id=f"{workflow_id}@{(idx + 1):03d}",
            step_id=rec.step_id,
            status=rec.status,
            summary={"completed_steps": idx + 1},
            artifact_refs=rec.artifact_refs,
        )
        for idx, rec in enumerate(records)
    ]
    snapshot = WorkflowState(
        workflow_id=workflow_id,
        input=None,
        policy={},
        limits={},
        plan=plan,
        status=step_results[-1].status if step_results else WorkflowStatus.PLANNED,
        steps=list(step_results),
        provenance=StepHistory(records=records),
        checkpoints=checkpoints,
        current_checkpoint=checkpoints[-1].checkpoint_id if checkpoints else None,
    )
    write_checkpoint(snapshot, workflow_id, base_dir)


def recover_workflow(workflow_id: str, base_dir: Optional[str] = None, resume_step_index: Optional[int] = None) -> "RecoveryResult":
    """Load, validate and classify a persisted workflow state (Phase C).

    Returns a RecoveryResult. Corrupted/incompatible state is never accepted;
    it is reported as classification INVALID_SCHEMA_OR_STATE with the error.

    When *resume_step_index* is given and a valid snapshot exists, the returned
    state's step_results reflect all steps before that index as completed, and
    the workflow is marked NEEDS_RECOVERY (recovery continues only after the
    operator confirms the step was safe to re-run).
    """
    from engine.mind_070.checkpoint import classify_recovery, recover_workflow as _recover

    if base_dir is None:
        base_dir = ".memind"
    result = _recover(workflow_id, base_dir=base_dir, resume_step_index=resume_step_index)
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


def _extract_workflow_error(result: StepResult) -> Optional["WorkflowError"]:
    output = result.output
    if isinstance(output, WorkflowError):
        return output
    if isinstance(output, dict) and "error" in output:
        return WorkflowError(str(output.get("error", "")), category=WorkflowErrorCategory(output.get("category", "EXECUTION_FAILURE")), step=None, retryable=False)
    return None
