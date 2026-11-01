"""Engine adapter protocol (Phase B binding placeholder).

Phase A defines the protocol. Phase B binds concrete adapters for MIND-010..MIND-060.

No engine-specific import belongs here so that the orchestrator stays
engine-standalone for this phase.
"""

from __future__ import annotations

from typing import Any, Protocol, runtime_checkable

from engine.mind_070.contracts import Step, StepResult, StepStatus


@runtime_checkable
class EngineAdapter(Protocol):
    """Protocol implemented by every engine adapter.

    ``execute(context, step)`` is a single synchronous call from the workflow
    executor returning a StepResult that downstream steps, retry policy, and
    recovery can reason about.
    """

    def execute(self, context: Any, step: Step) -> StepResult:
        ...
