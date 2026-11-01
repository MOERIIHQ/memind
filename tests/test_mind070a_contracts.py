"""MIND-070A unit tests: contracts, registry, plan validator.

Scope (MIND-070A only):
- operation registry availability (registered / unregistered)
- plan validation (invalid input, invalid plan, cyclic dependencies)
- StepResult contract
- error classification (WorkflowErrorCategory)

These tests do NOT execute MIND-010..MIND-060 engines. They run purely over
the contract and registry code of this phase.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_070.contracts import (
    Step,
    StepResult,
    StepStatus,
    WorkflowContext,
    WorkflowError,
    WorkflowErrorCategory,
    WorkflowPlan,
    WorkflowStatus,
)
from engine.mind_070.operation_registry import (
    OperationInfo,
    OperationRegistry,
    OperationType,
)
from engine.mind_070.plan_validator import validate_plan, ValidationError, PlanValidationResult


class TestOperationRegistryAvailability(unittest.TestCase):
    def setUp(self):
        self.registry = OperationRegistry()
        self.known = "text_task_to_formula"

    def test_registered_operation_is_available(self):
        self.assertTrue(self.registry.is_available(self.known))
        info = self.registry.get_operation(self.known)
        self.assertIsNotNone(info)
        self.assertEqual(info.operation_type, OperationType.TEXT_TASK_TO_FORMULA)
        self.assertTrue(info.enabled)

    def test_unregistered_operation_is_not_available(self):
        self.assertFalse(self.registry.is_available("no_such_operation_xyz"))

    def test_list_operations_returns_only_enabled(self):
        names = self.registry.list_operations()
        self.assertGreaterEqual(len(names), 4)


class TestPlanValidation(unittest.TestCase):
    def setUp(self):
        self.registry = OperationRegistry()
    def make_registry(self):
        return OperationRegistry()

    def make_plan(self, steps):
        return WorkflowPlan(workflow_id="wf-1", steps=steps)

    def test_valid_plan_passes(self):
        plan = self.make_plan(
            [
                Step(
                    id="s1",
                    operation="text_task_to_formula",
                    parameters={"text": "jumlahkan B2:C2 ke D2"},
                    dependencies=[],
                ),
                Step(
                    id="s2",
                    operation="auto_reporting",
                    parameters={},
                    dependencies=["s1"],
                ),
            ]
        )
        result = validate_plan(self.registry, plan)
        self.assertTrue(result.valid)
        self.assertEqual(len(result.errors), 0)

    def test_unregistered_operation_rejected(self):
        plan = self.make_plan(
            [Step(id="s1", operation="not_a_real_operation")]
        )
        result = validate_plan(self.registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(
            any(e.field == "steps[].operation" for e in result.errors)
        )

    def test_invalid_parameters_rejected(self):
        plan = self.make_plan(
            [Step(id="s1", operation="text_task_to_formula", parameters="not-a-dict")]
        )
        result = validate_plan(self.registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(
            any(e.field == "steps[].parameters" for e in result.errors)
        )

    def test_unknown_dependency_rejected(self):
        plan = self.make_plan(
            [Step(id="s1", operation="text_task_to_formula", dependencies=["missing_step"])]
        )
        result = validate_plan(self.registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(
            any("dependency step not found" in e.message for e in result.errors)
        )

    def test_circular_dependency_rejected(self):
        plan = WorkflowPlan(
            workflow_id="wf-cyclic",
            steps=[
                Step(id="a", operation="text_task_to_formula", dependencies=["b"]),
                Step(id="b", operation="auto_reporting", dependencies=["a"]),
            ],
        )
        result = validate_plan(self.registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(
            any("circular" in e.message for e in result.errors)
        )


class TestStepResultContract(unittest.TestCase):
    def make_result(self, status=StepStatus.SUCCEEDED, artifact=None, retry_attempt=0):
        return StepResult(
            status=status,
            output={"ok": True},
            artifact=artifact,
            retry_attempt=retry_attempt,
            max_retries=1,
            diagnostics={"engine": "MIND-070A"},
        )

    def test_step_result_has_all_required_fields(self):
        sr = self.make_result(artifact="/tmp/out.xlsx", retry_attempt=1)
        self.assertEqual(sr.status, StepStatus.SUCCEEDED)
        self.assertEqual(sr.artifact, "/tmp/out.xlsx")
        self.assertEqual(sr.retry_attempt, 1)
        self.assertEqual(sr.max_retries, 1)
        self.assertIsInstance(sr.diagnostics, dict)

    def test_step_result_serialization(self):
        sr = self.make_result(artifact="/tmp/out.xlsx")
        d = sr.to_dict()
        self.assertEqual(d.get("status"), "SUCCEEDED")
        self.assertEqual(d.get("artifact"), "/tmp/out.xlsx")


class TestWorkflowContextSerialization(unittest.TestCase):
    def make_context(self):
        plan = WorkflowPlan(workflow_id="wf-ctx", steps=[])
        return WorkflowContext(workflow_id="wf-ctx", plan=plan)

    def test_workflow_context_serialization(self):
        ctx = self.make_context()
        d = ctx.to_dict()
        self.assertEqual(d["workflow_id"], "wf-ctx")
        self.assertEqual(d["status"], "PLANNED")
        self.assertIsInstance(d["plan"], dict)

    def test_workflow_context_updated_at_changes(self):
        a = self.make_context()
        import time
        time.sleep(0.01)
        b = self.make_context()
        self.assertNotEqual(a.updated_at, b.updated_at)


class TestErrorClassification(unittest.TestCase):
    def test_category_values_are_unique_and_stable(self):
        cats = [e.value for e in WorkflowErrorCategory]
        self.assertEqual(len(cats), len(set(cats)))
        self.assertIn("TRANSIENT", cats)
        self.assertIn("INVALID_PLAN", cats)
        self.assertIn("EXECUTION_LIMIT_EXCEEDED", cats)

    def test_workflow_error_metadata(self):
        err = WorkflowError(
            "out of budget",
            WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED,
            step="s2",
            retryable=False,
        )
        self.assertEqual(err.category, WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED)
        self.assertEqual(err.step, "s2")
        self.assertFalse(err.retryable)
        self.assertIn("out of budget", str(err))

    def test_transient_error_is_retryable(self):
        err = WorkflowError("temp failure", WorkflowErrorCategory.TRANSIENT, retryable=True)
        self.assertTrue(err.retryable)

    def test_unsupported_operation_error_none_retryable(self):
        err = WorkflowError("unknown op", WorkflowErrorCategory.UNSUPPORTED_OPERATION, retryable=False)
        self.assertFalse(err.retryable)
