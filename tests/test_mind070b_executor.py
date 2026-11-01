"""Phase B (MIND-070B) executor + adapter tests.

Covers, at minimum:
- execution of a step with a registered operation,
- rejection of unregistered operations and invalid plans,
- dependency validation and step execution order,
- error propagation and classification,
- step result and required-output validation,
- mid-workflow failure without a false success,
- adapter compatibility with the actual MIND-010..060 component contracts.

Phase B does NOT implement state/provenance, full retry, limits, or cancellation
(see the parent package docstring of engine/mind_070/executor.py).
"""

import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES = os.path.join(ROOT, "fixtures")

from engine.mind_070.contracts import Step, StepStatus, WorkflowContext, WorkflowPlan, WorkflowError
from engine.mind_070.executor import execute_step, run_workflow, WorkflowResult
from engine.mind_070.operation_adapter_registry import build_phase_b_registry
from engine.mind_070.operation_registry import OperationContract, OperationInfo, OperationType, OperationRegistry
from engine.mind_070.plan_validator import validate_plan

try:
    import openpyxl  # openpyxl is used only to build a temporary workbook here
    HAVE_OPENPYXL = True
except ImportError:  # pragma: no cover
    HAVE_OPENPYXL = False


def _ctx(registry):
    plan = WorkflowPlan(workflow_id="wf-test")
    ctx = WorkflowContext(workflow_id="wf-test", plan=plan)
    ctx.registry = registry
    return ctx


def _out_path(tmpdir, name="out.xlsx"):
    return str(Path(tmpdir) / name)


def _auth_status(step_result):
    return step_result.status


class TestRegisteredOperationExecution(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_file_enhance_step(self):
        tmp = tempfile.mkdtemp()
        try:
            out_path = _out_path(tmp, "clean_enhanced.xlsx")
            step = Step(
                id="s1",
                operation="file_enhance",
                parameters={
                    "input_path": os.path.join(FIXTURES, "clean.xlsx"),
                    "output_path": out_path,
                },
            )
            result = execute_step(_ctx(self.registry), step)
            self.assertEqual(result.status, StepStatus.SUCCEEDED)
            self.assertTrue(result.output["valid"])
            self.assertEqual(result.artifact, out_path)
            self.assertEqual(result.output["issues_found"], 0)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_formula_apply_step_runs_component(self):
        tmp = tempfile.mkdtemp()
        try:
            if not HAVE_OPENPYXL:  # pragma: no cover
                self.skipTest("openpyxl not available")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(["A", "B", "C", "D"])
            ws.append([10, 20, 30, None])
            ws.append([40, 50, 60, None])
            in_path = os.path.join(tmp, "in.xlsx")
            out_path = _out_path(tmp, "out.xlsx")
            wb.save(in_path)

            step = Step(
                id="s1",
                operation="formula_apply",
                parameters={
                    "input_path": in_path,
                    "output_path": out_path,
                    "operation": {"target": "D2", "operation": "SUM", "operands": {"range": "A2:C2"}},
                },
            )
            result = execute_step(_ctx(self.registry), step)
            self.assertEqual(result.status, StepStatus.SUCCEEDED)
            self.assertTrue(result.output["formula_persisted"])
            self.assertEqual(result.output["formula_applied"], "=SUM(A2:C2)")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class TestUnregisteredOperationAndInvalidPlan(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_unregistered_operation_rejected(self):
        step = Step(id="s1", operation="no_such_operation_xyz", parameters={})
        result = execute_step(_ctx(self.registry), step)
        self.assertEqual(result.status, StepStatus.FAILED)
        self.assertEqual(result.diagnostics["category"], "UNSUPPORTED_OPERATION")

    def test_plan_with_unknown_dependency_rejected(self):
        plan = WorkflowPlan(
            workflow_id="wf-bad",
            steps=[Step(id="a", operation="file_enhance", dependencies=["missing_step"])],
        )
        with self.assertRaises(WorkflowError) as ctx:
            run_workflow("wf-bad", plan, registry=self.registry)
        self.assertEqual(ctx.exception.category, "INVALID_PLAN")

    def test_cyclic_plan_rejected(self):
        plan = WorkflowPlan(
            workflow_id="wf-cycle",
            steps=[
                Step(id="a", operation="file_enhance", dependencies=["b"]),
                Step(id="b", operation="file_enhance", dependencies=["a"]),
            ],
        )
        result = validate_plan(self.registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("circular" in e.message for e in result.errors))

    def test_disabled_operation_rejected(self):
        registry = OperationRegistry()
        registry.register_operation(
            OperationInfo(
                name="file_enhance",
                operation_type=OperationType.FILE_ENHANCE,
                contract=OperationContract(input={"i": "str"}, output={}),
                adapter=None,
                enabled=False,
            )
        )
        step = Step(id="s1", operation="file_enhance", parameters={})
        result = execute_step(_ctx(registry), step)
        self.assertEqual(result.status, StepStatus.FAILED)
        self.assertEqual(result.diagnostics["category"], "UNSUPPORTED_OPERATION")


class TestDependencyValidationAndExecutionOrder(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_order_respects_dependencies_not_list_order(self):
        tmp = tempfile.mkdtemp()
        try:
            a_out = _out_path(tmp, "a.xlsx")
            b_out = _out_path(tmp, "b.xlsx")
            # plan lists b first but b depends on a
            plan = WorkflowPlan(
                workflow_id="wf-order",
                steps=[
                    Step(id="b", operation="file_enhance", dependencies=["a"], parameters={
                        "input_path": os.path.join(FIXTURES, "clean.xlsx"), "output_path": b_out}),
                    Step(id="a", operation="file_enhance", dependencies=[], parameters={
                        "input_path": os.path.join(FIXTURES, "clean.xlsx"), "output_path": a_out}),
                ],
            )
            result = run_workflow("wf-order", plan, registry=self.registry)
            self.assertEqual(result.overall_status, StepStatus.SUCCEEDED)
            # execution order follows dependencies: a before b
            self.assertEqual(result.executed_step_ids, ["a", "b"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_step_without_registered_adapter_raises(self):
        registry = build_phase_b_registry()
        registry.register_operation(
            OperationInfo(
                name="formula_apply",
                operation_type=OperationType.FORMULA_APPLY,
                contract=OperationContract(input={"i": "str"}, output={}),
                adapter=None,
                enabled=True,
            )
        )
        plan = WorkflowPlan(
            workflow_id="wf-no-adapter",
            steps=[Step(id="s1", operation="formula_apply", parameters={})],
        )
        with self.assertRaises(WorkflowError) as ctx:
            run_workflow("wf-no-adapter", plan, registry=registry)
        self.assertEqual(ctx.exception.category, "UNSUPPORTED_OPERATION")


class TestErrorPropagationAndClassification(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_missing_input_file_classifies_as_execution_failure(self):
        tmp = tempfile.mkdtemp()
        try:
            missing = os.path.join(tmp, "does_not_exist.xlsx")
            step = Step(
                id="s1",
                operation="file_enhance",
                parameters={"input_path": missing, "output_path": _out_path(tmp, "o.xlsx")},
            )
            result = execute_step(_ctx(self.registry), step)
            self.assertEqual(result.status, StepStatus.FAILED)
            self.assertEqual(result.diagnostics["category"], "EXECUTION_FAILURE")
            self.assertIn("Cannot enhance workbook", str(result.output.get("error", "")))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_formula_on_invalid_cell_classifies_as_execution_failure(self):
        tmp = tempfile.mkdtemp()
        try:
            if not HAVE_OPENPYXL:  # pragma: no cover
                self.skipTest("openpyxl not available")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append([1, 2, 3, None])
            in_path = os.path.join(tmp, "in.xlsx")
            out_path = _out_path(tmp, "out.xlsx")
            wb.save(in_path)
            step = Step(
                id="s1",
                operation="formula_apply",
                parameters={
                    "input_path": in_path,
                    "output_path": out_path,
                    "operation": {"target": "ZZZ", "operation": "SUM", "operands": {"range": "A1:A1"}},
                },
            )
            result = execute_step(_ctx(self.registry), step)
            self.assertEqual(result.status, StepStatus.FAILED)
            self.assertEqual(result.diagnostics["category"], "EXECUTION_FAILURE")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_task_parse_error_classifies_as_validation_failure(self):
        step = Step(
            id="s1",
            operation="text_task_to_formula",
            parameters={"text": "", "input_path": "/tmp/x.xlsx", "output_path": "/tmp/y.xlsx"},
        )
        result = execute_step(_ctx(build_phase_b_registry()), step)
        self.assertEqual(result.status, StepStatus.FAILED)
        self.assertEqual(result.diagnostics["category"], "VALIDATION_FAILED")
        self.assertEqual(result.diagnostics.get("kind"), "INVALID")


class TestMidpointFailureDoesNotReportFalseSuccess(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_mid_step_failure_stops_and_not_succeeded(self):
        tmp = tempfile.mkdtemp()
        try:
            good_out = _out_path(tmp, "good.xlsx")
            missing = os.path.join(tmp, "missing.xlsx")
            s1 = Step(
                id="a",
                operation="file_enhance",
                parameters={"input_path": os.path.join(FIXTURES, "clean.xlsx"), "output_path": good_out},
            )
            s2 = Step(
                id="b",
                operation="file_enhance",
                parameters={"input_path": missing, "output_path": _out_path(tmp, "o.xlsx")},
            )
            plan = WorkflowPlan(workflow_id="wf-mid", steps=[s1, s2])
            result = run_workflow("wf-mid", plan, registry=self.registry)
            self.assertEqual(result.overall_status, StepStatus.FAILED)
            self.assertEqual(len(result.step_results), 2)
            self.assertEqual(result.step_results[0].status, StepStatus.SUCCEEDED)
            self.assertEqual(result.step_results[1].status, StepStatus.FAILED)
            self.assertIsNotNone(result.error)
            self.assertFalse(result.overall_status == StepStatus.SUCCEEDED)  # no false success
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_empty_plan_is_not_succeeded(self):
        plan = WorkflowPlan(workflow_id="wf-empty", steps=[])
        result = run_workflow("wf-empty", plan, registry=self.registry)
        self.assertEqual(result.overall_status, StepStatus.FAILED)
        self.assertIsNotNone(result.error)


class TestStepResultAndRequiredOutputValidation(unittest.TestCase):
    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_step_result_has_required_contract_fields(self):
        tmp = tempfile.mkdtemp()
        try:
            step = Step(
                id="s1",
                operation="formula_apply",
                parameters={"input_path": os.path.join(FIXTURES, "clean.xlsx"),
                            "output_path": _out_path(tmp, "o.xlsx"),
                            "operation": {"target": "D2", "operation": "SUM", "operands": {"range": "A2:A10"}}},
            )
            result = execute_step(_ctx(self.registry), step)
            self.assertEqual(result.status, StepStatus.SUCCEEDED)
            self.assertEqual(result.retry_attempt, 0)
            self.assertEqual(result.max_retries, 1)
            self.assertIsInstance(result.diagnostics, dict)
            self.assertIsInstance(result.to_dict(), dict)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class TestIntegrationAdapterCompatibility(unittest.TestCase):
    """End-to-end chains exercising real MIND-010..060 contracts."""

    def setUp(self):
        self.registry = build_phase_b_registry()

    def test_file_enhance_then_auto_report_not_required_chain(self):
        # file_enhance -> spreadsheet_intelligence -> auto_reporting chain
        tmp = tempfile.mkdtemp()
        try:
            enhanced = _out_path(tmp, "enhanced.xlsx")
            analysis_out = _out_path(tmp, "analysis.xlsx")  # not needed; analyze is read-only
            s1 = Step(
                id="enh",
                operation="file_enhance",
                parameters={"input_path": os.path.join(FIXTURES, "clean.xlsx"), "output_path": enhanced},
            )
            s2 = Step(
                id="intel",
                operation="spreadsheet_intelligence",
                parameters={"input_path": enhanced},
            )
            # reporting needs the analysis from s2; $ref substitutes the output
            s3 = Step(
                id="report",
                operation="auto_reporting",
                parameters={"spreadsheet_analysis": {"$ref": "intel"}},
            )
            plan = WorkflowPlan(
                workflow_id="wf-chain",
                steps=[s1, s2, s3],
            )
            result = run_workflow("wf-chain", plan, registry=self.registry)
            self.assertEqual(result.overall_status, StepStatus.SUCCEEDED)
            self.assertEqual(len(result.step_results), 3)
            self.assertEqual(result.step_results[0].status, StepStatus.SUCCEEDED)
            self.assertEqual(result.step_results[1].status, StepStatus.SUCCEEDED)
            self.assertEqual(result.step_results[2].status, StepStatus.SUCCEEDED)
            # report output is a Report.to_dict() shape, not a success lie
            self.assertIsInstance(result.step_results[2].output, dict)
            self.assertIn("metadata", result.step_results[2].output)
            self.assertEqual(result.step_results[2].output["metadata"]["source_type"], "xlsx")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_auto_reporting_requires_spreadsheet_analysis(self):
        step = Step(
            id="s1",
            operation="auto_reporting",
            parameters={},  # no spreadsheet_analysis -> VALIDATION_FAILED
        )
        result = execute_step(_ctx(build_phase_b_registry()), step)
        self.assertEqual(result.status, StepStatus.FAILED)
        self.assertEqual(result.diagnostics["category"], "VALIDATION_FAILED")


if __name__ == "__main__":
    unittest.main()
