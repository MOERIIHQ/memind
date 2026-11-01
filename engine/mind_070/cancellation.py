"""MIND-070D cancellation.

Cancellation is COOPERATIVE and only at SAFE boundaries (never mid-operation).
The orchestrator checks the token between steps and at workflow boundaries;
it must NOT interrupt a step that is mid-file, mid-database-transaction, or
otherwise destructive to restart the workflow cleanly from checkpoint.

Threading note: single-process, synchronous execution; the token is an
in-memory flag. For distributed/cross-process cancellation a gateway provider
(e.g. message-bus) must set this flag.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Optional

from engine.mind_070.contracts import WorkflowError, WorkflowErrorCategory

__all__ = ["CancellationToken", "request_cancellation", "cancel_requested", "is_cancelling"]


@dataclass
class CancellationToken:
    """Cooperative cancellation handle, safe to share across one workflow."""

    cancelled: bool = False
    cancelled_at: Optional[str] = None

    def request(self) -> None:
        self.cancelled = True
        self.cancelled_at = _now_iso()

    def reset(self) -> None:
        self.cancelled = False
        self.cancelled_at = None

    def is_set(self) -> bool:
        return self.cancelled


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def request_cancellation(token: Optional[CancellationToken]) -> None:
    if token is not None:
        token.request()


def cancel_requested(token: Optional[CancellationToken]) -> bool:
    return token is not None and token.cancelled


def is_cancelling(token: Optional[CancellationToken]) -> bool:
    return cancel_requested(token)
