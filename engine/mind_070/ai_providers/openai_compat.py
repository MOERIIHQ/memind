"""MIND-070F - Concrete AIProvider transport (OpenAI-compatible, stdlib only).

Status: Phase 2 transport adapter.  Implements the provider-agnostic
``AIProvider`` contract over an HTTP(S) endpoint using only the Python
standard library (``urllib.request``, ``ssl``, ``json``).

Architectural guard (D2/E):
- This adapter ONLY sends a request and returns a response.
- It NEVER runs an operation adapter, writes a workbook, changes workflow
  state, publishes artefacts, or runs PlanValidator/executor.
- The orchestration flow remains:
    AIProvider adapter -> candidate plan -> PlanValidator ->
    deterministic executor
- The PlanValidator/executor side is the single authority that may mutate
  state or exercise operations.  The adapter is a read-only transport.

Configuration (no credentials are stored in source, fixture, README, state or
provenance):
- ``AI_BASE_URL``   - HTTPS endpoint base URI (e.g. "https://api.example.com/v1")
- ``AI_API_KEY``    - bearer token (read from environment only; never logged)
- ``AI_MODEL``      - model name
- ``AI_TIMEOUT``    - request timeout in seconds (default 30 s maximum)
- ``AI_MAX_RESPONSE_BYTES`` - raw response cap (default 1 MiB)
- ``AI_MAX_PROMPT_BYTES``   - prompt cap (default 32 KiB UTF-8)
- ``AI_CHAT_ENDPOINT``      - path suffix (default "/chat/completions")

Retry policy (D3):
- At most one retry after the initial attempt (2 attempts total).
- Retryable (transient-safe): HTTP 429, HTTP 5xx, HTTP 408, network/timeout /
  connection-abort.
- NOT retryable (config/auth/permanent): 400/401/403/404/413/422, oversized
  response, invalid JSON, non-JSON content type.

No new third-party dependencies are required: the adapter is implemented with
pure ``urllib.request`` + ``ssl`` + ``json``.  Network access is only made when
the configured provider is present (AI_BASE_URL / AI_MODEL / AI_CHAT_ENDPOINT).
"""

from __future__ import annotations

import json
import os
import socket
import ssl
import urllib.error
import urllib.request

from typing import Any, Optional

from engine.mind_070.ai_provider import (
    AIError,
    AIProvider,
    AIProviderConfig,
    AICategory,
    DeterministicWorkflow,
)

__all__ = ["OpenAICompatProvider"]

# Default provisional size/time bounds (D4).  The transport enforces:
# raw response <= 1 MiB, prompt <= 32 KiB UTF-8.
_MAX_RESPONSE_BYTES = 1024 * 1024
_MAX_PROMPT_BYTES = 32 * 1024

# HTTP status codes that may be retried once (transient / transient-safe).
_RETRYABLE_HTTP = {429, 500, 502, 503, 504, 408}


# ---------------------------------------------------------------------------
# classification helpers
# ---------------------------------------------------------------------------

def _category_from_http_code(code: int) -> str:
    if code == 429:
        return AICategory.TIMEOUT
    if 500 <= code <= 504:
        return AICategory.TRANSIENT
    if code == 408:
        return AICategory.TIMEOUT
    if code in (401, 403):
        return AICategory.AUTH
    return AICategory.CONFIG


def _is_transient_http_code(code: int) -> bool:
    return code in _RETRYABLE_HTTP


def _is_transient_network(reason: BaseException) -> bool:
    return isinstance(reason, (socket.timeout, TimeoutError, ConnectionResetError, ConnectionAbortedError))


def _validate_config(timeout: float, max_response_bytes: int, max_prompt_bytes: int) -> None:
    if timeout <= 0:
        raise AIError(f"timeout must be > 0, got {timeout}", category=AICategory.CONFIG, retryable=False)
    if max_response_bytes < 0:
        raise AIError(f"max_response_bytes must be >= 0, got {max_response_bytes}", category=AICategory.CONFIG, retryable=False)
    if max_prompt_bytes < 0:
        raise AIError(f"max_prompt_bytes must be >= 0, got {max_prompt_bytes}", category=AICategory.CONFIG, retryable=False)


def _clean(value: Optional[str]) -> str:
    """Strip whitespace so a bare keyword or accidental space cannot break the
    endpoint / model / chat endpoint configuration."""
    return value.strip() if isinstance(value, str) else ""


# ---------------------------------------------------------------------------
# concrete transport adapter
# ---------------------------------------------------------------------------

class OpenAICompatProvider(AIProvider):
    """Concrete OpenAI-compatible provider over HTTPS.

    Provider-agnostic contract is supplied by the *interface* (AIProvider,
    AIProviderConfig, AIError, DeterministicWorkflow).  This class is a single
    concrete adapter.

    Responsibilities:
    - Send a minimal prompt to the configured endpoint and return the raw text.
    - Enforce timeout, raw-response-size, prompt-size and JSON candidate bounds.
    - Classify config/auth/transient errors and retry only transient-safe
      failures (at most one retry).

    Non-responsibilities (D2/E invariant):
    - Never run operation adapters, write workbooks, mutate workflow state,
      publish artefacts, or run PlanValidator/executor.
    """

    def __init__(
        self,
        config: Optional[AIProviderConfig] = None,
        *,
        base_url: Optional[str] = None,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        timeout: Optional[float] = None,
        max_response_bytes: Optional[int] = None,
        max_prompt_bytes: Optional[int] = None,
        chat_endpoint: Optional[str] = None,
    ) -> None:
        super().__init__(config)
        self.base_url = _clean(base_url if base_url is not None else os.environ.get("AI_BASE_URL", ""))
        self.api_key = api_key if api_key is not None else os.environ.get("AI_API_KEY", "")
        self.model = _clean(model if model is not None else os.environ.get("AI_MODEL", ""))
        self.timeout = float(timeout if timeout is not None else os.environ.get("AI_TIMEOUT", "30"))
        self.max_response_bytes = int(max_response_bytes or int(os.environ.get("AI_MAX_RESPONSE_BYTES", _MAX_RESPONSE_BYTES)))
        self.max_prompt_bytes = int(max_prompt_bytes or int(os.environ.get("AI_MAX_PROMPT_BYTES", _MAX_PROMPT_BYTES)))
        self.chat_endpoint = _clean(chat_endpoint if chat_endpoint is not None else os.environ.get("AI_CHAT_ENDPOINT", "/chat/completions"))

        _validate_config(self.timeout, self.max_response_bytes, self.max_prompt_bytes)

        # Configured but incomplete provider does NOT enable AI.  The
        # orchestrator stays deterministic until an explicit grant is made.
        self._ready = bool(self.base_url and self.model and self.chat_endpoint.startswith("/"))
        self._scheme = self.base_url.split(":", 1)[0].lower() if self.base_url else ""

    # ---------------------------------------------------------- API surface --
    def generate_text(self, prompt: str, *, as_json: bool = False) -> Any:
        """Return provider text (or parsed JSON when *as_json*)."""
        if not self._ready:
            raise AIError(
                "AI provider is not configured via AI_BASE_URL / AI_MODEL / "
                "AI_CHAT_ENDPOINT; no network will be opened",
                category=AICategory.CONFIG,
                retryable=False,
            )
        text = self._fetch_text(prompt)
        if as_json:
            try:
                return json.loads(text)
            except (ValueError, UnicodeDecodeError) as exc:
                raise AIError(
                    f"provider returned invalid JSON to generate_text(as_json=True): {exc}",
                    category=AICategory.MALFORMED,
                    retryable=False,
                ) from exc
        return text

    def generate_candidate_plan(self, prompt: str) -> Any:
        """Return the provider output as a *candidate* plan object.

        This is the D2 entry point.  The returned object is a plan *candidate*
        and MUST be passed to PlanValidator before execution.
        """
        if not self._ready:
            raise AIError(
                "AI provider not configured; deterministic fallback only",
                category=AICategory.CONFIG,
                retryable=False,
            )
        text = self._fetch_text(prompt)
        try:
            obj = json.loads(text)
        except (ValueError, UnicodeDecodeError) as exc:
            raise AIError(
                f"AI response is not valid JSON (candidate plan rejected): {exc}",
                category=AICategory.MALFORMED,
                retryable=False,
            ) from exc
        if not isinstance(obj, dict):
            raise AIError(
                f"AI response is not a plan object ({type(obj).__name__}); rejecting",
                category=AICategory.MALFORMED,
                retryable=False,
            )
        return obj

    # ------------------------------------------------------- transport call ---
    def _fetch_text(self, prompt: str) -> str:
        if not self._ready:
            raise AIError(
                "AI provider is not configured; deterministic fallback only",
                category=AICategory.CONFIG,
                retryable=False,
            )
        if len(prompt.encode("utf-8")) > self.max_prompt_bytes:
            raise AIError(
                f"prompt exceeds max_prompt_bytes ({len(prompt.encode('utf-8'))} > {self.max_prompt_bytes})",
                category=AICategory.CONFIG,
                retryable=False,
            )
        if not self.base_url.lower().startswith("https://"):
            raise AIError(
                f"endpoint must use HTTPS, got scheme {self._scheme!r}",
                category=AICategory.CONFIG,
                retryable=False,
            )
        req, payload = self._build_request(prompt)
        raw = self._fetch_with_retry(req, payload)
        content_type = raw.get("content_type") or ""
        if "application/json" not in content_type.lower():
            raise AIError(
                f"unexpected content-type {content_type!r}; expected application/json (malformed)",
                category=AICategory.MALFORMED,
                retryable=False,
            )
        return raw["body"].decode("utf-8", errors="replace")

    def _fetch_with_retry(self, req, prompt: str) -> dict:
        attempts = 0
        while True:
            try:
                return self._raw_urlopen(req)
            except urllib.error.HTTPError as exc:
                if _is_transient_http_code(exc.code):
                    attempts += 1
                    if attempts > 1:
                        raise self._make_http_error(exc) from exc
                    continue
                raise self._make_http_error(exc) from exc
            except socket.timeout as exc:
                attempts += 1
                if attempts > 1:
                    raise self._make_error("timeout", str(exc)) from exc
                continue
            except TimeoutError as exc:
                attempts += 1
                if attempts > 1:
                    raise self._make_error("timeout", str(exc)) from exc
                continue
            except urllib.error.URLError as exc:
                if _is_transient_network(exc.reason):
                    attempts += 1
                    if attempts > 1:
                        raise self._make_error("network", str(exc)) from exc
                    continue
                raise self._make_error("network", str(exc)) from exc
            except AIError:
                # Transport-level errors (e.g. oversized response) are raised by
                # the adapter itself and must preserve their category.
                raise
            except Exception as exc:  # defensive
                raise AIError(
                    f"unexpected transport error: {type(exc).__name__}",
                    category=AICategory.EXCEPTION,
                    retryable=False,
                ) from exc

    def _raw_urlopen(self, req) -> dict:
        ctx = ssl.create_default_context()
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
        response = opener.open(req, timeout=self.timeout)
        content_type = response.getheader("Content-Type") or ""
        body = self._read_bounded(response)
        return {"body": body, "content_type": content_type}

    def _read_bounded(self, response) -> bytes:
        """Bounded body read.

        Content-Length is available for a cheap up-front fail-fast; the body is
        then read in bounded chunks.  Any overflow is MALFORMED and is NOT
        forwarded to the executor or parsed further.
        """
        if self.max_response_bytes >= 0:
            try:
                declared = int(response.getheader("Content-Length") or 0)
            except (TypeError, ValueError):
                declared = 0
            if declared > self.max_response_bytes:
                raise AIError(
                    f"response declares Content-Length {declared}, exceeding "
                    f"max_response_bytes ({self.max_response_bytes})",
                    category=AICategory.MALFORMED,
                    retryable=False,
                )
        content = bytearray()
        while True:
            chunk = response.read(65536)
            if not chunk:
                break
            content.extend(chunk)
            if len(content) > self.max_response_bytes:
                raise AIError(
                    f"response exceeds max_response_bytes ({len(content)} > "
                    f"{self.max_response_bytes})",
                    category=AICategory.MALFORMED,
                    retryable=False,
                )
        return bytes(content)

    def _build_request(self, prompt: str):
        payload = json.dumps(
            {
                "model": self.model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 0,
            },
            ensure_ascii=True,
        ).encode("utf-8")
        authorization = f"Bearer {self.api_key}" if self.api_key else ""
        headers = {
            "Authorization": authorization,
            "Content-Type": "application/json",
            "Accept-Encoding": "identity",
            "User-Agent": "MIND-070-Transport/OpenAICompat",
        }
        url = self.base_url.rstrip("/") + self.chat_endpoint
        req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
        return req, payload

    def _make_http_error(self, exc) -> AIError:
        return AIError(
            f"provider HTTP {exc.code}: {exc.reason}",
            category=_category_from_http_code(exc.code),
            retryable=_is_transient_http_code(exc.code),
        )

    def _make_error(self, kind: str, detail: str) -> AIError:
        if kind == "timeout":
            return AIError(f"request timed out after {self.timeout}s", category=AICategory.TIMEOUT, retryable=True)
        if kind == "network":
            return AIError(f"network failure: {detail}", category=AICategory.TRANSIENT, retryable=True)
        return AIError(f"AI transport failure: {detail}", category=AICategory.CONFIG, retryable=False)
