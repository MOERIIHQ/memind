"""MIND-070F - Phase 2 transport / adapter mock tests (network-free by default).

Covers the Phase-2 checklist (D5 + E):

1. valid provider response
2. timeout, retry once
3. auth error (401) - not retried
4. HTTP 4xx permanent - no retry
5. HTTP 5xx - retry once, then succeed
6. response > 1 MiB rejected
7. prompt > 32 KiB rejected without network
8. transient retry at most one time
9. permanent error without retry
10. credentials not leaked to log / sent in request body
11. invalid AI candidate rejected by PlanValidator
12. valid plan runs only through the deterministic executor
13. deterministic workflow works without AI
14. default suite performs no network

All checks are mocked; no HTTP request is made in the default suite.
"""

import builtins
import json
import os
import shutil
import socket
import sys
import tempfile
import unittest
from contextlib import contextmanager
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.mind_070.ai_providers.openai_compat import OpenAICompatProvider
from engine.mind_070.ai_provider import AICategory, AIError
from engine.mind_070.contracts import Step, WorkflowPlan, WorkflowStatus, StepStatus
from engine.mind_070.operation_adapter_registry import get_default_registry
from engine.mind_070.plan_validator import validate_plan

# 40 KiB UTF-8 text > 32 KiB prompt cap (D4).
_AI_LARGE_PROMPT = "A" * 40000

_FIXTURES = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "fixtures")


class _FakeResponse:
    """Minimal stand-in for an HTTP(S) response."""

    def __init__(self, body=b"{}", status=200, content_type="application/json",
                 headers=None):
        self._body = body
        self._status = status
        self._content_type = content_type
        self._headers = dict(headers or {})
        self._headers.setdefault("Content-Type", content_type)
        self.read_bytes = 0
        self._pos = 0

    def getcode(self):
        return self._status

    def getheader(self, name, default=None):
        return self._headers.get(name, default)

    def read(self, size=-1):
        if self._pos >= len(self._body):
            return b""
        remaining = len(self._body) - self._pos
        end = self._pos + (size if size != -1 and size < remaining else remaining)
        chunk = self._body[self._pos:end]
        self._pos = end
        self.read_bytes += len(chunk)
        return chunk


class FakeOpener:
    """Fake urllib opener whose ``open`` returns one outcome per call."""

    def __init__(self, outcomes):
        self._outcomes = outcomes
        self.calls = 0

    def open(self, req, timeout=None):
        self.calls += 1
        kind = self._outcomes[self.calls - 1]
        if isinstance(kind, Exception):
            raise kind
        return kind


@contextmanager
def _patch_transport(outcomes):
    """Mock the adapter's urlopen.  *outcomes* is the per-call sequence."""
    opener = FakeOpener(outcomes)
    with mock.patch(
        "engine.mind_070.ai_providers.openai_compat.urllib.request.build_opener",
        return_value=opener,
    ):
        yield opener


def _http_error(code, msg="err"):
    import urllib.error
    return urllib.error.HTTPError(
        "https://example.com/v1/chat/completions", code, msg, {}, None)


def _openai_cfg(**overrides):
    base = {
        "base_url": "https://api.example.com/v1",
        "model": "llama-compat",
        "api_key": "k-AfN-s3cr3t-k3y",
        "timeout": 5.0,
        "max_response_bytes": 1048576,
        "max_prompt_bytes": 32768,
        "chat_endpoint": "/chat/completions",
    }
    base.update(overrides)
    return base


def _file_enhance_plan(output_path):
    return WorkflowPlan(
        workflow_id="wf-dt",
        steps=[Step(id="s1", operation="file_enhance", parameters={
            "input_path": os.path.join(_FIXTURES, "clean.xlsx"),
            "output_path": output_path,
        })],
    )


# ---------------------------------------------------------------------------
class TestValidResponse(unittest.TestCase):
    def test_valid_provider_response_yields_candidate_plan(self):
        payload = json.dumps({
            "steps": [{"id": "s1", "operation": "file_enhance", "parameters": {}}]
        }).encode()
        response = _FakeResponse(payload, content_type="application/json")
        with _patch_transport([response]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            self.assertEqual(
                p.generate_candidate_plan("sample text")["steps"][0]["operation"],
                "file_enhance",
            )
        self.assertEqual(opener.calls, 1)

    def test_generate_candidate_plan_rejects_non_json_output(self):
        response = _FakeResponse(b"<html>not-json</html>", content_type="text/html")
        with _patch_transport([response]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("sample text")
            self.assertEqual(_ctx.exception.category, AICategory.MALFORMED)
        self.assertEqual(opener.calls, 1)

    def test_prompt_too_long_rejected_without_network(self):
        with _patch_transport([]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan(_AI_LARGE_PROMPT)
            self.assertEqual(_ctx.exception.category, AICategory.CONFIG)
        self.assertEqual(opener.calls, 0)


class TestTransientRetry(unittest.TestCase):
    def test_retry_once_then_succeed(self):
        outcomes = [
            _http_error(503),
            _FakeResponse(
                json.dumps({"steps": [{"id": "s2", "operation": "formula_apply"}]}).encode(),
                content_type="application/json",
            ),
        ]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            out = p.generate_candidate_plan("chain")
        self.assertEqual(opener.calls, 2)
        self.assertEqual(out["steps"][0]["operation"], "formula_apply")

    def test_timeout_then_retry_then_succeed(self):
        outcomes = [socket.timeout(), _FakeResponse(
            json.dumps({"steps": []}).encode())]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            out = p.generate_candidate_plan("x")
        self.assertEqual(opener.calls, 2)
        self.assertEqual(out["steps"], [])

    def test_timeout_without_retry_when_max_exceeded(self):
        outcomes = [socket.timeout(), socket.timeout()]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.TIMEOUT)
        self.assertEqual(opener.calls, 2)


class TestPermanentAndAuthErrors(unittest.TestCase):
    def test_auth_401_not_retried(self):
        outcomes = [_http_error(401)]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.AUTH)
        self.assertEqual(opener.calls, 1)

    def test_http_400_not_retried(self):
        outcomes = [_http_error(400)]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.CONFIG)
        self.assertEqual(opener.calls, 1)

    def test_http_503_retried_and_then_succeed(self):
        outcomes = [_http_error(503), _FakeResponse(
            json.dumps({"steps": []}).encode())]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            out = p.generate_candidate_plan("x")
        self.assertEqual(opener.calls, 2)
        self.assertEqual(out["steps"], [])

    def test_transient_503_gives_up_after_one_retry(self):
        outcomes = [_http_error(503), _http_error(503)]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.TRANSIENT)
        self.assertEqual(opener.calls, 2)

    def test_oversized_response_rejected(self):
        headers = {"Content-Length": "2097152"}  # 2 MiB > 1 MiB cap
        response = _FakeResponse(b"{}", content_type="application/json", headers=headers)
        with _patch_transport([response]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.MALFORMED)
        self.assertEqual(opener.calls, 1)

    def test_prompt_too_long_rejected_without_network(self):
        with _patch_transport([]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan(_AI_LARGE_PROMPT)
            self.assertEqual(_ctx.exception.category, AICategory.CONFIG)
        self.assertEqual(opener.calls, 0)


class TestCredentialsAndSecurity(unittest.TestCase):
    def test_credentials_not_sent_in_request_body(self):
        captured = {}
        original_build_request = OpenAICompatProvider._build_request

        def recording(self_ref, prompt):
            req, payload = original_build_request(self_ref, prompt)
            captured["payload"] = payload
            return req, payload

        with mock.patch.object(OpenAICompatProvider, "_build_request", recording):
            with _patch_transport([_FakeResponse(b"{}", content_type="application/json")]) as opener:
                p = OpenAICompatProvider(**_openai_cfg())
                p.generate_candidate_plan("x")
        self.assertNotIn(b"k-AfN-s3cr3t-k3y", captured["payload"])
        self.assertEqual(opener.calls, 1)

    def test_credentials_not_logged(self):
        with _patch_transport([_http_error(401)]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with mock.patch("builtins.print") as fake_print:
                with self.assertRaises(AIError) as _ctx:
                    p.generate_candidate_plan("x")
                self.assertEqual(_ctx.exception.category, AICategory.AUTH)
                fake_print.assert_not_called()
        self.assertEqual(opener.calls, 1)


# ---------------------------------------------------------------------------
class TestPlanValidatorOrdering(unittest.TestCase):
    def test_invalid_ai_candidate_rejected_by_planvalidator(self):
        registry = get_default_registry()
        from engine.mind_070.contracts import Step, WorkflowPlan
        invalid_steps = [
            Step(id="s1", operation="file_enhance", parameters={}),
            Step(id="s1", operation="formula_apply", parameters={}),
        ]
        plan = WorkflowPlan(workflow_id="w", steps=invalid_steps)
        result = validate_plan(registry, plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("duplicate" in e.message for e in result.errors))

    def test_valid_plan_runs_via_deterministic_executor(self):
        from engine.mind_070.executor import run_workflow
        out = tempfile.mkdtemp()
        try:
            plan = _file_enhance_plan(os.path.join(out, "out.xlsx"))
            res = run_workflow("wf-dt", plan)
            self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
            self.assertEqual(res.step_results[0].status, StepStatus.SUCCEEDED)
        finally:
            shutil.rmtree(out, ignore_errors=True)

    def test_deterministic_workflow_runs_without_ai(self):
        from engine.mind_070.executor import run_workflow
        with _patch_transport([]) as opener:
            res = run_workflow(
                "wf-dt", _file_enhance_plan(tempfile.mktemp(suffix=".xlsx")),
            )
        self.assertEqual(res.overall_status, WorkflowStatus.SUCCEEDED)
        self.assertEqual(opener.calls, 0)  # no network when no AI provider

    # ------------------------------------------------------------------
    # Retry consistency (Phase 2.1)
    # ------------------------------------------------------------------

    def test_malformed_not_retryable_by_default(self):
        err = AIError("malformed payload", category=AICategory.MALFORMED)
        self.assertFalse(err.retryable)
        self.assertEqual(err.category, AICategory.MALFORMED)
        # A genuinely transient category keeps its retryability by default.
        ok = AIError("timeout", category=AICategory.TIMEOUT)
        self.assertTrue(ok.retryable)

    def test_json_malformed_rejected(self):
        response = _FakeResponse(b"{invalid json", content_type="application/json")
        with _patch_transport([response]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("sample text")
            self.assertEqual(_ctx.exception.category, AICategory.MALFORMED)
            self.assertFalse(_ctx.exception.retryable)
        self.assertEqual(opener.calls, 1)

    def test_oversized_response_not_retried(self):
        headers = {"Content-Length": "2097152"}  # 2 MiB > 1 MiB cap
        response = _FakeResponse(b"{}", content_type="application/json",
                                 headers=headers)
        with _patch_transport([response]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.MALFORMED)
            self.assertFalse(_ctx.exception.retryable)
        self.assertEqual(opener.calls, 1)

    def test_transient_retry_max_two_attempts_total(self):
        outcomes = [_http_error(503), _http_error(503)]
        with _patch_transport(outcomes) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.TRANSIENT)
        # Initial attempt + one retry = at most two attempts total.
        self.assertEqual(opener.calls, 2)

    def test_auth_error_not_retryable(self):
        err = AIError("unauthorized", category=AICategory.AUTH)
        self.assertFalse(err.retryable)
        with _patch_transport([_http_error(401)]) as opener:
            p = OpenAICompatProvider(**_openai_cfg())
            with self.assertRaises(AIError) as _ctx:
                p.generate_candidate_plan("x")
            self.assertEqual(_ctx.exception.category, AICategory.AUTH)
            self.assertFalse(_ctx.exception.retryable)
        self.assertEqual(opener.calls, 1)

    def test_default_suite_performs_no_network(self):
        # Force urlopen to fail on ANY call; the default MIND-070 suite must
        # still complete and make zero transport calls.
        with _patch_transport([socket.timeout()]) as opener:
            from engine.mind_070.executor import run_workflow
            plan = _file_enhance_plan(tempfile.mktemp(suffix=".xlsx"))
            res = run_workflow("wf-nw", plan)
        self.assertEqual(opener.calls, 0)
