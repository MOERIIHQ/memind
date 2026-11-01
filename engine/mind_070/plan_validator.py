"""Plan validation against the operation registry.

MIND-070A scope: ``validate_plan()`` establishes that a workflow plan only
references registered, enabled operations, has unique step ids, valid
parameter types, and no cyclic dependencies. It does NOT execute plans.

A validation result must be consumed by an executor that supplies its own
bounds; validation itself is intentionally pure (no side effects).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from engine.mind_070.contracts import Step, WorkflowPlan
from engine.mind_070.operation_registry import OperationRegistry


@dataclass(frozen=True)
class ValidationError:
    field: str
    message: str
    detail: Any = None


@dataclass
class PlanValidationResult:
    plan: WorkflowPlan
    valid: bool
    errors: list[ValidationError] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def add_error(self, field: str, message: str, detail: Any = None) -> None:
        self.errors.append(ValidationError(field, message, detail))


def validate_plan(registry: OperationRegistry, plan: WorkflowPlan) -> PlanValidationResult:
    """Return a read-only validation of *plan* against *registry*.

    A plan is valid when:
    - workflow_id is non-empty;
    - step ids are unique;
    - every step operation is registered and enabled;
    - every step parameters value is a dict;
    - every declared dependency step id exists in the plan;
    - the dependency graph has no cycles.
    """
    result = PlanValidationResult(plan=plan, valid=True)

    if not plan.workflow_id:
        result.add_error("workflow_id", "workflow_id must not be empty")
        result.valid = False

    plan_map = {step.id: step for step in plan.steps}
    if len(plan_map) != len(plan.steps):
        result.add_error("steps", "duplicate step id is not allowed")
        result.valid = False

    for step in plan.steps:
        _validate_step(result, step, registry, plan_map)

    _check_cycles(result, plan_map)

    if not plan.output_policy:
        result.warnings.append(
            "output_policy is not set; default new_file=True, overwrite=False is recommended"
        )
    if not plan.limits:
        result.warnings.append(
            "limits are not set; a workflow executor must still provide hard caps"
        )
    return result


def _validate_step(
    result: PlanValidationResult,
    step: Step,
    registry: OperationRegistry,
    plan_map: dict[str, Step],
) -> None:
    if not step.id:
        result.add_error("steps[].id", "step id must not be empty")
    if not step.operation:
        result.add_error("steps[].operation", "step operation must not be empty")
        return

    if not registry.is_available(step.operation):
        result.add_error(
            "steps[].operation",
            f"unregistered or disabled operation: {step.operation}",
            {"available": registry.list_operations()},
        )
        result.valid = False

    if not isinstance(step.parameters, dict):
        result.add_error("steps[].parameters", "step parameters must be a dict")
        result.valid = False

    for dep in step.dependencies:
        if dep not in plan_map:
            result.add_error(
                "steps[].dependencies",
                f"dependency step not found: {dep}",
            )
            result.valid = False


def _check_cycles(result: PlanValidationResult, plan_map: dict[str, Step]) -> None:
    visiting: set[str] = set()
    visited: set[str] = set()

    def visit(step_id: str) -> None:
        if step_id in visited:
            return
        if step_id in visiting:
            result.add_error(
                "steps[].dependencies",
                f"circular dependency detected involving step: {step_id}",
            )
            result.valid = False
            return
        step = plan_map.get(step_id)
        if step is None:
            return
        visiting.add(step_id)
        for dep in step.dependencies:
            visit(dep)
        visiting.discard(step_id)
        visited.add(step_id)

    for step_id in list(plan_map.keys()):
        visit(step_id)
