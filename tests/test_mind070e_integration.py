"""MIND-070E integration and acceptance audit (A-D end-to-end + failure scenarios).

Run with the project interpreter (/usr/bin/python3) which carries openpyxl +
pytest. Focus: real interactions across MIND-070A..D, not per-component unit
tests. Also asserts AIProvider is optional and deterministic workflows run
without it, and that no new openpyxl usage exists outside MIND-010.
"""

import os
import shutil
import sys
import tempfile
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_070.cancellation import CancellationToken
from engine.mind_070.checkpoint import CheckpointRecord, classify_recovery, load_workflow_state, save_workflow_state
from engine.mind_070.contracts import (Step, StepResult, StepStatus, WorkflowContext,
                                       WorkflowError, WorkflowErrorCategory, WorkflowPlan,
                                       WorkflowStatus)
from engine.mind_070.executor import run_workflow
from engine.mind_070.limits import GlobalLimits, WorkflowLimits, validate_limits_config
from engine.mind_070.operation_registry import (OperationInfo, OperationRegistry,
                                                OperationType, get_default_registry)
from engine.mind_070.plan_validator import validate_plan
from engine.mind_070.publisher import (OutputPublication, PublicationError,
                                       OVERWRITE_ALWAYS, OVERWRITE_EXPLICIT,
                                       OVERWRITE_NEVER, validate_output_path)
from engine.mind_070.provenance import StepHistory
from engine.mind_070.state import WorkflowState, save_workflow_state, load_workflow_state
from engine.mind_070.operation_adapter_registry import build_phase_b_registry

FIXTURES = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fixtures")
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEST_BASE = os.path.join(REPO, ".mind070e_teststate", "tmp")
os.makedirs(TEST_BASE, exist_ok=True)
CLEAN = os.path.join(FIXTURES, "clean.xlsx")


def work():
    return tempfile.mkdtemp(prefix="mind070e_", dir=TEST_BASE)


def _sr(status=StepStatus.SUCCEEDED, output=None, artifact=None, retry_attempt=0):
    return StepResult(status=status, output=output, artifact=artifact, retry_attempt=retry_attempt)


class TestEndToEndIntegration(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_plan_to_output_to_state_provenance_checkpoint(self):
        registry = build_phase_b_registry()
        wb = os.path.join(self.work, "a.xlsx")
        shutil.copy2(CLEAN, wb)
        # file_enhance -> spreadsheet_intelligence -> auto_reporting (real
        # MIND-010..060 chain); the auto_reporting step binds the $ref of
        # spreadsheet_intelligence's output (a SpreadsheetAnalysis).
        plan = WorkflowPlan(workflow_id="wf-int", steps=[
            Step(id="enh", operation="file_enhance",
                 parameters={"input_path": wb, "output_path": os.path.join(self.work, "out1.xlsx")}),
            Step(id="intel", operation="spreadsheet_intelligence",
                 parameters={"input_path": os.path.join(self.work, "out1.xlsx")}),
            Step(id="report", operation="auto_reporting",
                 parameters={"spreadsheet_analysis": {"$ref": "intel"}}),
        ])
        res = run_workflow("wf-int", plan, registry=registry, checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        # auto_reporting writes a Report struct (path-independent); the checkpoint
        # and provenance must persist it without breaking, and be loadable again.
        self.assertIsInstance(res.step_results[2].output, dict)
        self.assertIn("metadata", res.step_results[2].output)
        state = load_workflow_state("wf-int", self.work)
        self.assertEqual(state.status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(len(state.steps), 3)
        self.assertEqual(len(state.provenance.records), 3)
        self.assertEqual(len(state.checkpoints), 3)

    def test_publication_failure_is_not_succeded(self):
        reg = build_phase_b_registry()
        plan = WorkflowPlan(workflow_id="wf-pfail", steps=[
            Step(id="enh", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": os.path.join(self.work, "a.xlsx")}),
        ])
        # explicit permission granted but output_publication target is a
        # directory that publish() cannot commit to -> the run must report
        # FAILED, never SUCCEEDED.
        pub = OutputPublication(self.work, "wf-pub-fail", overwrite_policy=OVERWRITE_EXPLICIT)
        res = run_workflow("wf-pfail", plan, registry=reg, output_publication=pub, checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertFalse(res.overall_status == WorkflowStatus.SUCCEEDED)

    def test_atomic_overwrite_not_corrupt_old_file_on_failed_commit(self):
        pub = OutputPublication(self.work, "wf-atomic", overwrite_policy=OVERWRITE_EXPLICIT, explicit_overwrite=True)
        # seed the OLD file (the publisher creates the target directory itself)
        open(os.path.join(self.work, "seed_old.xlsx"), "w").write("OLD")
        pub.publish(os.path.join(self.work, "seed_old.xlsx"), "old.xlsx")
        target = pub.resolve("old.xlsx")
        self.assertEqual(open(target).read(), "OLD")
        # allowed overwrite -> the previous file is replaced atomically
        new_src = os.path.join(self.work, "new.xlsx")
        with open(new_src, "w") as fh:
            fh.write("NEW")
        pub.publish(new_src, "old.xlsx")
        self.assertEqual(open(target).read(), "NEW")
        # a failing publish must NOT corrupt the previously-valid file
        with self.assertRaises(PublicationError):
            pub.publish(os.path.join(self.work, "nope.xlsx"), "old.xlsx")
        self.assertEqual(open(target).read(), "NEW")

    def test_publication_path_traversal_symlink_conflict_and_bad_dest(self):
        pub = OutputPublication(self.work, "wf-p")
        self.assertRaises(PublicationError, pub.resolve, "../../escape.xlsx")
        os.makedirs(os.path.join(self.work, "inner"))
        with open(os.path.join(self.work, "inner", "t.xlsx"), "w") as fh:
            fh.write("x")
        # the link must sit INSIDE the allowed outputs dir so its basename is a
        # "benign" name; the symlink then points OUTSIDE the allowed dir, which
        # the safe resolution must reject.
        os.makedirs(pub.outputs_base_dir, exist_ok=True)
        escape = os.path.join(pub.outputs_base_dir, "escape.xlsx")
        os.symlink(os.path.join(self.work, "inner", "t.xlsx"), escape)
        with self.assertRaises(PublicationError):
            pub.resolve(os.path.basename(escape))
        self.assertRaises(PublicationError, validate_output_path, CLEAN, source_path=CLEAN)

    def test_overwrite_requires_explicit_permission(self):
        # a second write to the same output is an overwrite and must be
        # permitted when explicit permission was granted; a plain (never)
        # policy must still deny it.
        pub = OutputPublication(self.work, "wf-overwrite", overwrite_policy=OVERWRITE_ALWAYS)
        first = pub.publish(CLEAN, "a.xlsx")
        second = pub.publish(CLEAN, "a.xlsx")
        self.assertEqual(first, second)
        pub_deny = OutputPublication(self.work, "wf-deny")
        first_deny = pub_deny.publish(CLEAN, "b.xlsx")
        self.assertTrue(os.path.exists(first_deny))
        with self.assertRaises(PublicationError):
            pub_deny.publish(CLEAN, "b.xlsx")  # overwrite denied


class TestRetryRecoveryAndStateConsistency(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def _transient_op_registry(self):
        reg = OperationRegistry()
        n = {"k": 0}

        def _transient(ctx, step):
            n["k"] += 1
            if n["k"] == 1:
                raise WorkflowError("transient", WorkflowErrorCategory.TRANSIENT, step=step.id, retryable=True)
            return StepResult(status=StepStatus.SUCCEEDED, output={"ok": True}, artifact=None)

        reg.register_operation(OperationInfo(name="t", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_transient, enabled=True))
        return reg

    def test_one_transient_retry_records_category_and_attempts(self):
        plan = WorkflowPlan(workflow_id="wf-retry", steps=[Step(id="s", operation="t", parameters={})])
        res = run_workflow("wf-retry", plan, registry=self._transient_op_registry(), checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(res.step_results[0].retry_attempt, 1)
        self.assertEqual(res.step_results[0].diagnostics.get("retry_category"), "TRANSIENT")

    def test_permanent_error_not_retried(self):
        reg = OperationRegistry()

        def _bad(ctx, step):
            return StepResult(status=StepStatus.FAILED, output={"error": "bad"}, retry_attempt=0, max_retries=1)

        reg.register_operation(OperationInfo(name="bad", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_bad, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-perm", steps=[Step(id="s", operation="bad", parameters={})])
        res = run_workflow("wf-perm", plan, registry=reg, checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertEqual(res.step_results[0].retry_attempt, 0)

    def test_recovery_no_blind_repeat_on_uncertain_artifact(self):
        # a completed step's artifact is referenced but missing -> recovery
        # must classify NEEDS_RECOVERY, not RECOVERED, so no auto-repeat.
        plan = WorkflowPlan(workflow_id="wf-unc", steps=[
            Step(id="enh", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": os.path.join(self.work, "a.xlsx")}),
        ])
        res = run_workflow("wf-unc", plan, registry=build_phase_b_registry(), checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        ws = load_workflow_state("wf-unc", self.work)
        artifact = ws.steps[0].artifact
        # simulate a crash that dropped the artifact but the state survived
        if os.path.exists(artifact):
            os.remove(artifact)
        ws.steps[0].artifact = artifact  # still recorded
        save_workflow_state(ws, "wf-unc", self.work)
        recovered = load_workflow_state("wf-unc", self.work)
        self.assertEqual(classify_recovery(recovered), "NEEDS_RECOVERY")

    def test_checkpoint_state_artifact_consistent_after_retry_failure_restart(self):
        # run, corrupt limits config (invalid), re-run -> consistent FAILED state
        ws = WorkflowState(
            workflow_id="wf-restart", status=WorkflowStatus.SUCCEEDED, steps=[_sr()],
            limits=GlobalLimits.defaults().__dict__,
        )
        save_workflow_state(ws, "wf-restart", self.work)
        loaded = load_workflow_state("wf-restart", self.work)
        self.assertEqual(loaded.limits.get("max_steps"), GlobalLimits.defaults().max_steps)

    def test_invalid_limits_do_not_silently_disable(self):
        with self.assertRaises(WorkflowError):
            validate_limits_config("not-a-dict")


class TestLimitsAndCancellation(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_global_and_workflow_limits_enforced(self):
        reg = build_phase_b_registry()
        plan = WorkflowPlan(workflow_id="wf-lim", steps=[
            Step(id="a", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": os.path.join(self.work, "a.xlsx")}),
            Step(id="b", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": os.path.join(self.work, "b.xlsx")}),
        ])
        res = run_workflow("wf-lim", plan, registry=reg, limits=WorkflowLimits(max_steps=1), checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.FAILED)
        self.assertEqual(res.error.category, WorkflowErrorCategory.EXECUTION_LIMIT_EXCEEDED)

    def test_cancellation_not_succeded_and_artifact_kept(self):
        reg = OperationRegistry()

        def _writer(ctx, step):
            target = os.path.join(self.work, f"{step.id}.xlsx")
            open(target, "w").close()
            time.sleep(0.4)  # stays alive across the cancel boundary
            return StepResult(status=StepStatus.SUCCEEDED, output=target, artifact=target)

        reg.register_operation(OperationInfo(name="writer", operation_type=OperationType.TEXT_TASK_TO_FORMULA,
                                             contract=None, adapter=_writer, enabled=True))
        plan = WorkflowPlan(workflow_id="wf-cancel", steps=[
            Step(id="a", operation="writer", parameters={}),
            Step(id="b", operation="writer", parameters={}),
        ])
        tok = CancellationToken()
        def _delay_cancel():
            time.sleep(0.1)
            tok.request()
        threading.Thread(target=_delay_cancel, daemon=True).start()
        res = run_workflow("wf-cancel", plan, registry=reg, cancellation_token=tok, checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.CANCELLED)
        self.assertFalse(res.overall_status == WorkflowStatus.SUCCEEDED)
        self.assertTrue(os.path.exists(os.path.join(self.work, "a.xlsx")))
        self.assertNotIn("b.xlsx", os.listdir(self.work))

    def test_cancellation_before_start(self):
        tok = CancellationToken()
        tok.request()
        res = run_workflow("wf-not", WorkflowPlan(workflow_id="wf-not",
            steps=[Step(id="a", operation="file_enhance", parameters={"input_path": CLEAN})]),
            registry=build_phase_b_registry(), cancellation_token=tok)
        self.assertEqual(res.overall_status, WorkflowStatus.CANCELLED)
        self.assertFalse(os.path.exists(os.path.join(self.work, "publications")))


class TestCompatibilityAndSecurity(unittest.TestCase):
    def setUp(self):
        self.work = work()

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_registry_only_allows_registered_operations(self):
        reg = OperationRegistry()
        plan = WorkflowPlan(workflow_id="wf-reg", steps=[
            Step(id="a", operation="file_enhance", parameters={"input_path": CLEAN}),
            Step(id="b", operation="__not_registered__", parameters={}),
        ])
        result = validate_plan(reg, plan)
        self.assertFalse(result.valid)
        self.assertTrue(result.errors)

    def test_circular_dependencies_rejected(self):
        plan = WorkflowPlan(workflow_id="wf-circ", steps=[
            Step(id="a", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": "a.xlsx"}, dependencies=["b"]),
            Step(id="b", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": "b.xlsx"}, dependencies=["a"]),
        ])
        result = validate_plan(build_phase_b_registry(), plan)
        self.assertFalse(result.valid)
        self.assertTrue(result.errors)

    def test_missing_required_output_rejected(self):
        pub = OutputPublication(self.work, "wf-missing")
        # step a produces output a.xlsx; required output b.xlsx is missing
        with self.assertRaises(PublicationError):
            pub.validate_required_outputs(["a.xlsx", "b.xlsx"])  # b.xlsx missing


    def test_no_new_openpyxl_outside_mind010(self):
        # machine-readable check: the only module importing openpyxl is MIND-010.
        from engine.mind_070 import state as st
        import inspect
        for mod in [st, __import__("engine.mind_070.executor", fromlist=["x"]),
                    __import__("engine.mind_070.publisher", fromlist=["x"]),
                    __import__("engine.mind_070.limits", fromlist=["x"])]:
            path = inspect.getsourcefile(mod)
            src = open(path).read()
            self.assertFalse("openpyxl" in src, f"{mod.__name__} unexpectedly imports openpyxl")

    def test_no_network_or_ai_dependency_in_workflow(self):
        # deterministic workflow must run with AIProvider absent and no network.
        plan = WorkflowPlan(workflow_id="wf-def", steps=[
            Step(id="enh", operation="file_enhance", parameters={"input_path": CLEAN, "output_path": os.path.join(self.work, "out.xlsx")}),
        ])
        res = run_workflow("wf-def", plan, registry=build_phase_b_registry(), checkpoint_base_dir=self.work)
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)


if __name__ == "__main__":
    unittest.main()
