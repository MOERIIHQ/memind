"""MIND-070E - Provider-agnostic AIProvider foundation (Phase 1).

Status: contract + validation foundation.  No network calls, no vendor SDK and
no concrete provider.  This module only defines the interface, the config
validation rules and the deterministic-fallback mechanism.

The orchestrator stays deterministic by construction: when no AIProvider or
when ``deterministic_fallback=True`` is set, every workflow runs on the
rule-based engine (MIND-010..MIND-060) exactly as before.  Any AI call is a
thin, optional overlay and must never be required for a workflow to complete.

Network rule (D5):
- No SDK, transport or connection is registered here.
- A provider is NOT automatically usable just because a configuration object
  exists.  Activation requires an explicit agent/UI grant that a concrete
  provider adapter is installed and safe to use.
- During Phase 1, a configured provider is never connected.  Wiring of the
  transport to a real network is deferred to a later phase after a concrete
  provider and network adapter are approved.
- No credentials are read from, printed to, or persisted by this module.

Retry separation (D3):
- Workflow retry and AI retry are separate budgets.  This module only defines
  the classification of errors; the workflow executor owns the workflow retry
  budget and never consumes it for an AI request.
"""

from __future__ import annotations

__all__ = [
    "AIProvider",
    "AIProviderConfig",
    "AIError",
    "DeterministicWorkflow",
    # Error classification (D1).
    "AICategory",
    "classify_ai_error",
]

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Error classification (D1): distinguish config, auth, timeout, transient,
# malformed and permanent failures so they are NOT all treated as retryable.
# ---------------------------------------------------------------------------

class AICategory:
    """Stable category strings for AI errors.  Only categories listed as
    retryable are safe to re-issue automatically."""
    CONFIG = "AI_CONFIG"        # mis-configured provider/setup (permanent)
    AUTH = "AI_AUTH"            # invalid/expired credentials (permanent)
    TIMEOUT = "AI_TIMEOUT"      # exceeded the requested timeout (transient)
    TRANSIENT = "AI_TRANSIENT"  # transient provider failure (transient)
    MALFORMED = "AI_MALFORMED"  # unparsable/invalid payload (transient)
    EXCEPTION = "AI_EXCEPTION"  # unexpected error the provider cannot classify
    PERMANENT = "AI_PERMANENT"  # wrong model/operation foreseen to fail (permanent)


# retryable is a statement about the AI layer only; the workflow runner keeps
# the workflow retry budget separate and never inherits it.
_RETRYABLE_CATEGORIES = {
    AICategory.TIMEOUT,
    AICategory.TRANSIENT,
}
_NON_RETRYABLE_CATEGORIES = {
    AICategory.CONFIG,
    AICategory.AUTH,
    AICategory.PERMANENT,
}


class AIError(Exception):
    """Base error for the optional AI layer.  Category-aware so the orchestrator
    can decide retry (transient) vs. halt (permanent)."""

    def __init__(
        self,
        message: str,
        category: str,
        *,
        retryable: Optional[bool] = None,
        diagnostics: Optional[dict[str, Any]] = None,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.category = category
        self.retryable = retryable if retryable is not None else category in _RETRYABLE_CATEGORIES
        self.diagnostics = diagnostics or {}


def classify_ai_error(
    exception: Optional[BaseException] = None,
    *,
    category: Optional[str] = None,
    message: Optional[str] = None,
    retryable_override: Optional[bool] = None,
) -> tuple[str, bool]:
    """Return ``(category, retryable)`` for an AI error without assuming every
    error is retryable.

    Rules:
    - status/shape/class errors -> CONFIG (permanent),
    - credential / auth errors -> AUTH (permanent),
    - time-out / malformed payload -> MALFORMED or TIMEOUT (transient),
    - unknown transient failure -> TRANSIENT (transient),
    - any explicitly permanent failure -> PERMANENT (permanent).

    ``classify_ai_error(...)`` is used by the AIProvider implementation to map
    provider signals into categories before they reach the workflow runner.
    """
    if exception is not None and isinstance(exception, AIError):
        exc = exception
        category = exc.category if category is None else category
        retryable = exc.retryable if retryable_override is None else retryable_override
    else:
        category = category or AICategory.CONFIG
        retryable = retryable_override if retryable_override is not None else (
            category in _RETRYABLE_CATEGORIES
        )
        message = message or (str(exception) if exception is not None else "")

    # The caller may override; otherwise we derive from the category alone.
    if retryable_override is not None:
        retryable = retryable_override
    return category if category is not None else AICategory.CONFIG, retryable


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

@dataclass
class AIProviderConfig:
    """Configuration surface for an optional AI provider.

    Contract notes:
    - Config validation is performed by :meth:`validate` before any provider
      instance is used.  Invalid config is a permanent error and NEVER triggers
      a network connection.
    - ``deterministic_fallback=True`` means the orchestrator MUST skip any AI
      call and run deterministically.
    - A provider is NOT automatically usable just because a config object
      exists.  Integration status is tracked separately (see module docstring).
    """

    provider: Optional[str] = None
    model: Optional[str] = None
    timeout_seconds: float = 30.0      # maximum 30 s per request (provisional)
    temperature: float = 0.0
    max_tokens: int = 0
    # When True the orchestrator MUST run deterministically and skip any AI call.
    deterministic_fallback: bool = True

    # ------------------------------------------------------------------------
    # Provisional size/time bounds (D4).  These remain configurable and are
    # enforced at the AIProvider layer (later phase); here they are documented
    # policy so the runner never reads an unbounded provider.
    # ------------------------------------------------------------------------
    MAX_TIMEOUT_SECONDS = 30.0        # per-request timeout (provisional)
    MAX_RESPONSE_BYTES = 1024 * 1024  # AI response (provisional, 1 MiB)
    MAX_CANDIDATE_STEPS = 120         # candidate plan steps (provisional)
    MAX_CANDIDATE_JSON_BYTES = 1024 * 1024  # candidate JSON (provisional, 1 MiB)
    MAX_PROMPT_BYTES = 32 * 1024        # AI prompt (provisional, 32 KiB)

    def validate(self) -> None:
        """Validate config values.  Raises :class:`AIConfigError` if invalid.

        Validation is VALUE-level only (ranges, required naming).  It does not
        assert that a provider is installed/available -- that is a separate
        integration-status check and must never be conflated with config
        validity.
        """
        from .ai_provider import AICategory, AIError

        if self.timeout_seconds < 0:
            raise AIError(
                f"timeout_seconds must be >= 0, got {self.timeout_seconds}",
                category=AICategory.CONFIG,
                retryable=False,
            )
        if self.temperature < 0 or self.temperature > 2.0:
            raise AIError(
                f"temperature must be in [0, 2], got {self.temperature}",
                category=AICategory.CONFIG,
                retryable=False,
            )
        if self.max_tokens < 0:
            raise AIError(
                f"max_tokens must be >= 0, got {self.max_tokens}",
                category=AICategory.CONFIG,
                retryable=False,
            )
        if self.deterministic_fallback and (self.provider or self.model):
            # Deterministic mode takes precedence; provider/model must not be set.
            raise AIError(
                "deterministic_fallback=True with provider/model set; "
                "deterministic_fallback must be False to use a provider",
                category=AICategory.CONFIG,
                retryable=False,
            )

    def to_dict(self) -> dict[str, Any]:
        return {
            "provider": self.provider,
            "model": self.model,
            "timeout_seconds": self.timeout_seconds,
            "temperature": self.temperature,
            "max_tokens": self.max_tokens,
            "deterministic_fallback": self.deterministic_fallback,
        }


# ---------------------------------------------------------------------------
# Provider contract
# ---------------------------------------------------------------------------

class AIProvider(ABC):
    """Minimum interface an optional AI provider must satisfy.

    Responsibilities (D1):
    - produce CANDIDATE workflow plans, never execute operations,
    - respect the timeout and response-size bounds,
    - classify errors clearly (config/auth/timeout/transient/malformed/
      permanent), NEVER marking every error retryable,
    - not store or print credentials, and not persist workbook payloads.

    Identity rule: a configured provider is NOT usable until an explicit
    integration grant (separate from configuration) marks it available.
    """

    def __init__(self, config: Optional[AIProviderConfig] = None) -> None:
        self.config = config or AIProviderConfig()
        self.config.validate()  # fail fast on invalid configuration

    @abstractmethod
    def generate_text(self, prompt: str, *, as_json: bool = False) -> Any:
        """Return text (or parsed JSON when as_json=True) that a *candidate*
        workflow plan; never an executed side effect.

        Raises :class:`AIError` with an explicit category on:
        - config/auth problems (permanent),
        - time-out (transient),
        - malformed output (transient),
        - other transient/permanent failures.

        The provider must not store credentials or workbook payloads and must
        not perform network calls unless a later phase wires the transport.
        """

    def validate_json(self, payload: Any) -> bool:
        """Return True only if *payload* is well-formed output from this
        provider.  Non-conforming output is rejected, never accepted."""
        raise NotImplementedError

    # ------------------------------------------------------------------
    # Response-size guard (D4).  Enforcement is deferred to the AIProvider
    # implementation; this module does not enforce or measure it.
    # ------------------------------------------------------------------
    # (placeholder for future provider implementations; not part of the
    #  current freeze and no network is established here.)




# ---------------------------------------------------------------------------
# Deterministic fallback (unchanged semantics)
# ---------------------------------------------------------------------------

class DeterministicWorkflow:
    """Flag and helper indicating the orchestrator is running without AI."""

    def __init__(self, deterministic: bool = True) -> None:
        self.deterministic = deterministic

    @classmethod
    def is_deterministic(cls, provider: Optional["AIProvider"]) -> bool:
        return provider is None or provider.config.deterministic_fallback
