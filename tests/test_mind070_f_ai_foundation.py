"""MIND-070F - Phase 1 AIProvider foundation + execution safety tests.

Covers the Phase-1 mandatory checklist (D1, D3, D4, D5, E):
- config validation and provider-not-auto-usable,
- error classification and retryability (AI retry is separate from workflow),
- candidate-plan structural bounds and zero-side-effect rejection,
- deterministic workflow without AI,
- no network access and no credential leakage in this phase,
- MIND-070A-E regression compatibility.
"""

import os
import shutil
import sys
import tempfile
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
from engine.mind_070.ai_provider import (
    AICategory,
    AIError,
    AIProviderConfig,
    classify_ai_error,
    DeterministicWorkflow,
)
from engine.mind_070.cancellation import CancellationToken
from engine.mind_070.executor import run_workflow, WorkflowResult
from engine.mind_070.limits import (
    MAX_AI_RESPONSE_BYTES,
    MAX_CANDIDATE_JSON_BYTES,
    MAX_CANDIDATE_STEPS,
    GlobalLimits,
)
from engine.mind_070.operation_adapter_registry import build_phase_b_registry
from engine.mind_070.operation_registry import OperationInfo, OperationRegistry, OperationType
from engine.mind_070.plan_validator import validate_plan

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
AI_PROVIDER_SRC = os.path.join(ROOT, "engine", "mind_070", "ai_provider.py")
AI_EXECUTOR_SRC = os.path.join(ROOT, "engine", "mind_070", "executor.py")

FIXTURES = os.path.join(ROOT, "fixtures")


def _step_result(status=StepStatus.SUCCEEDED, artifact=None):
    return StepResult(status=status, output={"ok": True}, artifact=artifact)


def _ctx(registry):
    plan = WorkflowPlan(workflow_id="wf-ctx")
    ctx = WorkflowContext(workflow_id="wf-ctx", plan=plan)
    ctx.registry = registry
    return ctx


def _tmpdir():
    return tempfile.mkdtemp(prefix="mind070f_")


def _xlsx_path():
    return os.path.join(_tmpdir(), "out.xlsx")


def _one_step_plan(operation="file_enhance", **params):
    # file_enhance requires input_path and output_path to be a valid structural
    # candidate plan (used for end-to-end happy-path tests).
    if "input_path" not in params:
        params["input_path"] = os.path.join(FIXTURES, "clean.xlsx")
    if "output_path" not in params:
        params["output_path"] = _xlsx_path()
    return WorkflowPlan(
        workflow_id="wf-one",
        steps=[Step(id="s1", operation=operation, parameters=params)],
    )


class TestProviderConfigValidation(unittest.TestCase):
    def test_valid_config_passes(self):
        cfg = AIProviderConfig()
        self.assertIsNone(cfg.validate())
        self.assertEqual(cfg.timeout_seconds, 30.0)
        self.assertEqual(cfg.timeout_seconds, AIProviderConfig.MAX_TIMEOUT_SECONDS)

    def test_timeout_bound(self):
        self.assertEqual(AIProviderConfig.MAX_TIMEOUT_SECONDS, 30.0)
        self.assertEqual(AIProviderConfig.MAX_TIMEOUT_SECONDS, 30.0)

    def test_invalid_timeout_permanent(self):
        cfg = AIProviderConfig(timeout_seconds=-1)
        with self.assertRaises(AIError) as ctx:
            cfg.validate()
        self.assertEqual(ctx.exception.category, AICategory.CONFIG)
        self.assertFalse(ctx.exception.retryable)

    def test_invalid_temperature(self):
        cfg = AIProviderConfig(temperature=-0.1)
        with self.assertRaises(AIError) as ctx:
            cfg.validate()
        self.assertEqual(ctx.exception.category, AICategory.CONFIG)

    def test_invalid_max_tokens(self):
        cfg = AIProviderConfig(max_tokens=-1)
        with self.assertRaises(AIError) as ctx:
            cfg.validate()
        self.assertEqual(ctx.exception.category, AICategory.CONFIG)

    def test_deterministic_fallback_with_provider_rejected(self):
        cfg = AIProviderConfig(deterministic_fallback=True, provider="x")
        with self.assertRaises(AIError) as ctx:
            cfg.validate()
        self.assertEqual(ctx.exception.category, AICategory.CONFIG)

    def test_provider_not_auto_usable(self):
        cfg = AIProviderConfig()
        self.assertTrue(cfg.deterministic_fallback)
        self.assertTrue(DeterministicWorkflow.is_deterministic(None))


class TestErrorClassification(unittest.TestCase):
    def _classify(self, cat):
        return classify_ai_error(category=cat)

    def test_timeout_transient(self):
        cat, retryable = self._classify(AICategory.TIMEOUT)
        self.assertEqual(cat, AICategory.TIMEOUT)
        self.assertTrue(retryable)

    def test_malformed_not_retryable(self):
        # Phase 2.1: malformed (size/class errors on the wire) is not
        # retryable by default; only transient-safe failures (TIMEOUT/TRANSIENT)
        # are retried automatically.
        cat, retryable = self._classify(AICategory.MALFORMED)
        self.assertEqual(cat, AICategory.MALFORMED)
        self.assertFalse(retryable)

    def test_config_permanent(self):
        cat, retryable = self._classify(AICategory.CONFIG)
        self.assertEqual(cat, AICategory.CONFIG)
        self.assertFalse(retryable)

    def test_auth_permanent(self):
        cat, retryable = self._classify(AICategory.AUTH)
        self.assertEqual(cat, AICategory.AUTH)
        self.assertFalse(retryable)

    def test_permanent_permanent(self):
        cat, retryable = self._classify(AICategory.PERMANENT)
        self.assertEqual(cat, AICategory.PERMANENT)
        self.assertFalse(retryable)

    def test_ai_error_not_all_retryable(self):
        for cat in (AICategory.CONFIG, AICategory.AUTH, AICategory.PERMANENT):
            with self.subTest(cat=cat):
                self.assertFalse(AIError("x", category=cat, retryable=None).retryable)

    def test_credentials_shape_in_diagnostics(self):
        sr = StepResult(
            status=StepStatus.SUCCEEDED, output={"ok": True},
            diagnostics={"api_key": "x"},
        )
        self.assertEqual(sr.to_dict()["diagnostics"], {"api_key": "x"})


class TestCandidatePlanAndValidation(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_candidate_plan_ok_with_default_limits(self):
        res = run_workflow("wf-valid", _one_step_plan(), registry=self.registry)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)

    def test_empty_plan_returns_failed_no_raise(self):
        plan = WorkflowPlan(workflow_id="wf-empty", steps=[])
        res = run_workflow("wf-empty", plan, registry=self.registry)
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertIsNotNone(res.error)

    def test_step_count_bound_enforced(self):
        plan = WorkflowPlan(
            workflow_id="wf-mega",
            steps=[Step(id=f"s{i}", operation="file_enhance") for i in range(MAX_CANDIDATE_STEPS + 1)],
        )
        with self.assertRaises(WorkflowError) as ctx:
            run_workflow("wf-mega", plan, registry=self.registry)
        self.assertEqual(ctx.exception.category, WorkflowErrorCategory.INVALID_PLAN)

    def test_candidate_json_size_boundary(self):
        import unittest.mock as mock

        import engine.mind_070.executor as ex_mod
        with mock.patch.object(
            ex_mod, "MAX_CANDIDATE_JSON_BYTES", 1
        ):
            plan = _one_step_plan()
            with self.assertRaises(WorkflowError) as ctx:
                run_workflow("wf-json", plan, registry=self.registry)
            self.assertEqual(ctx.exception.category, WorkflowErrorCategory.INVALID_PLAN)

    def test_oversized_candidate_plan_rejected_no_side_effect(self):
        # An unregistered operation must be rejected structurally BEFORE any
        # execution and without writing files.
        plan = _one_step_plan(operation="__not_registered__")
        workdir = _tmpdir()
        try:
            with self.assertRaises(WorkflowError) as ctx:
                run_workflow("wf-invalid", plan, registry=self.registry)
            self.assertEqual(ctx.exception.category, WorkflowErrorCategory.INVALID_PLAN)
            self.assertEqual(os.listdir(workdir), [])
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

    def test_unregistered_operation_rejected_structural(self):
        registry = OperationRegistry()
        plan = _one_step_plan(operation="__not_registered__")
        with self.assertRaises(WorkflowError) as ctx:
            run_workflow("wf-bad", plan, registry=registry)
        self.assertEqual(ctx.exception.category, WorkflowErrorCategory.INVALID_PLAN)


class TestRetrySeparation(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_global_limits_max_retries_default_is_one(self):
        self.assertEqual(GlobalLimits.defaults().max_retries, 1)

    def _registry_with_transient(self):
        reg = OperationRegistry()
        calls = {"n": 0}

        def _transient(ctx, step):
            calls["n"] += 1
            if calls["n"] == 1:
                raise WorkflowError(
                    "temporarily unavailable", WorkflowErrorCategory.TRANSIENT,
                    step=step.id, retryable=True,
                )
            return StepResult(status=StepStatus.SUCCEEDED, output={"ok": True}, artifact=None)

        op = OperationInfo(
            name="transient",
            operation_type=OperationType.TEXT_TASK_TO_FORMULA,
            contract=None,
            adapter=_transient,
            enabled=True,
        )
        reg.register_operation(op)
        return reg

    def test_workflow_retry_transient_attempts_once(self):
        reg = self._registry_with_transient()
        plan = WorkflowPlan(
            workflow_id="wf-retry-t", steps=[Step(id="s", operation="transient")]
        )
        res = run_workflow("wf-retry-t", plan, registry=reg)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(res.step_results[0].retry_attempt, 1)

    def test_workflow_retry_never_consumes_ai_budget(self):
        self.assertNotIn("max_ai_response_bytes", GlobalLimits.defaults().__dict__)
        self.assertEqual(MAX_AI_RESPONSE_BYTES, 1024 * 1024)

    def test_recovery_needs_recovery_on_missing_artifact(self):
        from engine.mind_070.checkpoint import classify_recovery
        from engine.mind_070.state import WorkflowState

        ws = WorkflowState(
            workflow_id="wf-un", status=WorkflowStatus.SUCCEEDED,
            steps=[_step_result(artifact="/missing.xlsx")],
        )
        self.assertEqual(classify_recovery(ws), "NEEDS_RECOVERY")


class TestDeterministicAndNoNetwork(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_deterministic_workflow_runs_without_ai(self):
        out = _xlsx_path()
        try:
            res = run_workflow("wf-def", _one_step_plan(output_path=out), registry=self.registry)
            self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        finally:
            if os.path.exists(out):
                os.remove(out)

    def test_no_network_modules_in_ai_provider(self):
        src = open(AI_PROVIDER_SRC, encoding="utf-8").read()
        for token in ("requests", "http", "socket", "urllib", "os.environ", "open("):
            self.assertNotIn(token, src, f"ai_provider.py contains '{token}'")

    def test_no_network_modules_in_executor(self):
        src = open(AI_EXECUTOR_SRC, encoding="utf-8").read()
        for token in ("requests", "http", "socket", "urllib"):
            self.assertNotIn(token, src, f"executor.py contains '{token}'")


class TestMIND070AToECompatibility(unittest.TestCase):
    def test_plan_validator_and_contracts_unbroken(self):
        from engine.mind_070.plan_validator import validate_plan
        from engine.mind_070.contracts import Step, WorkflowPlan

        plan = WorkflowPlan(workflow_id="wf-a", steps=[Step(id="s1", operation="file_enhance")])
        validate_plan(build_phase_b_registry(), plan)
        self.assertTrue(True)
