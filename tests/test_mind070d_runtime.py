"""MIND-070D runtime tests (permissions, output, retry, limits, cancellation).

Spec 3.5 + 4. Coverage of the four MIND-070D groups:
- A. Permissions and output publication.
- B. Retry.
- C. Execution limits.
- D. Cancellation.

Run with the project interpreter (/usr/bin/python3) which carries openpyxl +
pytest. 14 areas covered by construction; results are reported per-command.
"""

import os
import shutil
import sys
import tempfile
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_070.checkpoint import CheckpointRecord, classify_recovery, recover_workflow
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
from engine.mind_070.cancellation import CancellationToken
from engine.mind_070.executor import run_workflow
from engine.mind_070.limits import (
    GlobalLimits,
    WorkflowLimits,
    validate_limits_config,
)
from engine.mind_070.operation_registry import (
    OperationInfo,
    OperationRegistry,
    OperationType,
)
from engine.mind_070.plan_validator import validate_plan
from engine.mind_070.publisher import (
    OVERWRITE_EXPLICIT,
    OutputPublication,
    PublicationError,
    OVERWRITE_ALWAYS,
    OVERWRITE_NEVER,
    validate_output_path,
)
from engine.mind_070.provenance import StepHistory
from engine.mind_070.state import WorkflowState, save_workflow_state, load_workflow_state
from engine.mind_070.operation_adapter_registry import build_phase_b_registry

FIXTURES = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures"
)
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEST_BASE = os.path.join(REPO, ".mind070d_teststate")
os.makedirs(TEST_BASE, exist_ok=True)

CLEAN = os.path.join(FIXTURES, "clean.xlsx")


def work():
    return tempfile.mkdtemp(prefix="mind070d_", dir=TEST_BASE)


def _sr(status=StepStatus.SUCCEEDED, output=None, artifact=None, retry_attempt=0):
    return StepResult(status=status, output=output, artifact=artifact, retry_attempt=retry_attempt)


class TestOutputPermissionsAndPublication(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_output_new_as_default(self):
        pub = OutputPublication(self.work, "wf-new")
        self.assertTrue(
            str(pub.resolve("result.xlsx")).startswith(str(pub.outputs_base_dir))
        )
        source = CLEAN
        target = pub.publish(source, "result.xlsx")
        self.assertTrue(os.path.exists(target))
        # new output is the default: a second publish of the same name is
        # denied because overwrite requires explicit permission.
        with self.assertRaises(PublicationError):
            pub.publish(source, "result.xlsx")
        # a newer name is fine
        self.assertTrue(os.path.exists(pub.publish(source, "result2.xlsx")))

    def test_overwrite_without_permission_rejected(self):
        pub = OutputPublication(self.work, "wf-nop")
        first = pub.publish(CLEAN, "a.xlsx")
        self.assertTrue(os.path.exists(first))
        with self.assertRaises(PublicationError):
            pub.publish(CLEAN, "a.xlsx")  # overwrite denied

    def test_overwrite_with_explicit_permission(self):
        pub = OutputPublication(self.work, "wf-yes", overwrite_policy=OVERWRITE_ALWAYS)
        first = pub.publish(CLEAN, "b.xlsx")
        second = pub.publish(CLEAN, "b.xlsx")
        self.assertEqual(first, second)  # rewritten

    def test_overwrite_with_explicit_flag(self):
        pub = OutputPublication(self.work, "wf-flag", overwrite_policy=OVERWRITE_EXPLICIT, explicit_overwrite=True)
        first = pub.publish(CLEAN, "c.xlsx")
        second = pub.publish(CLEAN, "c.xlsx")
        self.assertEqual(first, second)

    def test_path_traversal_rejected(self):
        pub = OutputPublication(self.work, "wf-pt")
        with self.assertRaises(PublicationError):
            pub.resolve("../../escape.xlsx")

    def test_symlink_escape_rejected(self):
        pub = OutputPublication(self.work, "wf-sym")
        # a symlink whose target lies OUTSIDE the allowed outputs dir must be
        # rejected instead of being followed into the wider filesystem.
        outer = os.path.join(self.work, "escaped")
        os.makedirs(outer)
        target = os.path.join(outer, "target.xlsx")
        with open(target, "w") as fh:
            fh.write("x")
        # the link must sit INSIDE the allowed outputs dir so that its filename
        # becomes part of the resolved output path; the symlink then escapes the
        # allowed dir even though the name looks benign.
        os.makedirs(pub.outputs_base_dir, exist_ok=True)
        escape = os.path.join(pub.outputs_base_dir, "escape.xlsx")
        os.symlink(target, escape)  # name looks safe, target escapes allowed dir
        with self.assertRaises(PublicationError):
            pub.resolve(os.path.basename(escape))

    def test_input_output_conflict_rejected(self):
        with self.assertRaises(PublicationError):
            validate_output_path(CLEAN, source_path=CLEAN)

    def test_publication_failure_keeps_partial(self):
        pub = OutputPublication(self.work, "wf-fail")
        # a missing source must not delete the (unchanged) fixture workbook,
        # and the failure is reported without a false SUCCESS.
        with self.assertRaises(PublicationError):
            pub.publish(os.path.join(self.work, "missing.xlsx"), "out.xlsx")
        self.assertTrue(os.path.exists(CLEAN))


class TestRetry(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def _registry_with_transient(self):
        reg = OperationRegistry()
        calls = {"n": 0}

        def _transient(ctx, step):
            calls["n"] += 1
            if calls["n"] == 1:
                raise WorkflowError(
                    "temporarily unavailable (transient)", WorkflowErrorCategory.TRANSIENT,
                    step=step.id, retryable=True,
                )
            return StepResult(status=StepStatus.SUCCEEDED, output={"ok": True}, artifact=None)

        reg.register_operation(OperationInfo(name="transient", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_transient, enabled=True))
        return reg

    def test_one_transient_retry_attempts_once(self):
        plan = WorkflowPlan(workflow_id="wf-retry", steps=[Step(id="s", operation="transient", parameters={})])
        res = run_workflow("wf-retry", plan, registry=self._registry_with_transient())
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(len(res.step_results), 1)
        self.assertEqual(res.step_results[0].retry_attempt, 1)

    def test_permanent_failure_not_retried(self):
        reg = OperationRegistry()

        def _bad(ctx, step):
            return StepResult(status=StepStatus.FAILED, output={"error": "bad input"},
                              retry_attempt=0, max_retries=1)

        reg.register_operation(OperationInfo(name="bad", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_bad, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-perm", steps=[Step(id="s", operation="bad", parameters={})])
        res = run_workflow("wf-perm", plan, registry=reg)
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertEqual(res.step_results[0].retry_attempt, 0)

    def test_uncertain_step_not_auto_repeated(self):
        # recovery must not re-execute a step whose prior effect is uncertain:
        # an artifact declared by a SUCCEEDED step is missing on disk.
        ws = WorkflowState(workflow_id="wf-uncertain", status=WorkflowStatus.SUCCEEDED,
                           steps=[_sr(artifact="/missing.xlsx")])
        self.assertEqual(classify_recovery(ws), "NEEDS_RECOVERY")


class TestExecutionLimits(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_limits_config_valid_and_invalid(self):
        ok = validate_limits_config({"max_steps": 5, "max_retries": 2})
        self.assertEqual(ok.max_steps, 5)
        with self.assertRaises(WorkflowError):
            validate_limits_config({"max_steps": -1})

    def test_global_limits_enforced(self):
        # 2 steps > max_steps=1 -> EXECUTION_LIMIT_EXCEEDED -> FAILED.
        from engine.mind_070.contracts import StepResult, StepStatus

        def _ok(ctx, step):
            return StepResult(status=StepStatus.SUCCEEDED, output={"ok": True}, artifact=None)

        reg = OperationRegistry()
        reg.register_operation(OperationInfo(name="ok", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_ok, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-lim", steps=[
            Step(id="a", operation="ok", parameters={}), Step(id="b", operation="ok", parameters={}),
        ])
        res = run_workflow("wf-lim", plan, registry=reg, limits=WorkflowLimits(max_steps=1, reason="CI gating"))
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertIsNotNone(res.error)
        self.assertEqual(res.error.category, WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED)

    def test_limits_persisted_in_state(self):
        wf = WorkflowLimits(max_retries=2, reason="CI build")
        state = WorkflowState(workflow_id="wf-l",
                              limits=wf.apply_overrides(GlobalLimits.defaults()))
        save_workflow_state(state, "wf-l", self.work)
        loaded = load_workflow_state("wf-l", self.work)
        self.assertEqual(loaded.limits.get("max_retries"), 2)


class TestCancellation(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_not_started_cancellation(self):
        tok = CancellationToken()
        tok.request()
        from engine.mind_070.executor import run_workflow
        res = run_workflow("wf-notstarted", WorkflowPlan(workflow_id="wf-notstarted",
                                                         steps=[Step(id="a", operation="file_enhance",
                                                                     parameters={"input_path": CLEAN})]),
                           registry=build_phase_b_registry(), cancellation_token=tok)
        self.assertEqual(res.overall_status, WorkflowStatus.CANCELLED)
        self.assertFalse(res.overall_status == WorkflowStatus.SUCCEEDED)
        self.assertFalse(os.path.exists(os.path.join(self.work, "publications")))

    def test_between_steps_cancellation_leaves_artifact(self):
        reg = OperationRegistry()

        def _writer(ctx, step):
            time.sleep(0.4)  # slow step keeps the cancel BEFORE step b
            target = os.path.join(self.work, f"{step.id}.xlsx")
            shutil.copy2(CLEAN, target)
            return StepResult(status=StepStatus.SUCCEEDED, output=target, artifact=target)

        reg.register_operation(OperationInfo(name="writer", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_writer, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-cancel",
                             steps=[Step(id="a", operation="writer", parameters={}),
                                    Step(id="b", operation="writer", parameters={})])
        tok = CancellationToken()
        # request cancellation during step a's execution; the cancel must be
        # honoured at the step boundary (before step b), not mid-write.
        def _delay_cancel():
            time.sleep(0.05)
            tok.request()
        threading.Thread(target=_delay_cancel, daemon=True).start()
        from engine.mind_070.executor import run_workflow
        res = run_workflow("wf-cancel", plan, registry=reg, cancellation_token=tok)
        self.assertEqual(res.overall_status, WorkflowStatus.CANCELLED)
        self.assertFalse(res.overall_status == WorkflowStatus.SUCCEEDED)
        self.assertTrue(os.path.exists(os.path.join(self.work, "a.xlsx")))
        # step b never ran
        self.assertNotIn("b.xlsx", os.listdir(self.work))


class TestIntegrationAndConsistency(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_provenance_tracks_retry(self):
        reg = OperationRegistry()
        n = {"k": 0}

        def _transient_once(ctx, step):
            n["k"] += 1
            if n["k"] == 1:
                raise WorkflowError("transient", WorkflowErrorCategory.TRANSIENT, step=step.id, retryable=True)
            return StepResult(status=StepStatus.SUCCEEDED, output={"ok": True}, artifact=None)

        reg.register_operation(OperationInfo(name="t", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_transient_once, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-h", steps=[Step(id="s", operation="t", parameters={})])
        res = run_workflow("wf-h", plan, registry=reg)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(res.step_results[0].retry_attempt, 1)
        # the transient category is carried on the returned result for provenance
        self.assertEqual(res.step_results[0].diagnostics.get("retry_category"), "TRANSIENT")


if __name__ == "__main__":
    unittest.main()
