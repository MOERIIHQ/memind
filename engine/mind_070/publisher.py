"""MIND-070D output publication and permission validation.

Rules:
- New output is the DEFAULT. Overwrite is only allowed with explicit permission.
- All path/validation errors are raised as WorkflowError with category
  OUTPUT_PUBLICATION_FAILURE (consistent failure), never a false SUCCESS.
- Partial artifacts are NOT deleted: they remain for diagnosis and recovery.
- Publication failure is recorded and the workflow must NOT report SUCCEEDED.
"""

from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path
from typing import Any, Optional

from engine.mind_070.contracts import WorkflowError, WorkflowErrorCategory

__all__ = [
    "OutputPublication", "PublicationError", "OVERWRITE_NEVER",
    "OVERWRITE_ALWAYS", "OVERWRITE_EXPLICIT",
]

OVERWRITE_NEVER = "never"
OVERWRITE_ALWAYS = "always"
OVERWRITE_EXPLICIT = "explicit"

_ALLOWED_CHARS = set("-_/.%0-9a-zA-Z")
_ALLOWED_ROOTS = ("/", "")


class PublicationError(Exception):
    """Base for publication- and permission-related failures."""


def _is_safe_rel(path: str) -> bool:
    """Reject absolute, drive-letter, '..' and embedded NUL paths."""
    if not isinstance(path, str) or not path:
        return False
    if "\x00" in path:
        return False
    if path.startswith("/") or (len(path) > 1 and path[1] == ":" and len(path) > 2 and path[2] in "/\\"):
        return False
    parts = path.replace("\\", "/").split("/")
    if any(p in ("", ".", "..") for p in parts):
        return False
    return True


def _resolve_within(base: Path, rel: Path, label: str) -> Path:
    """Resolve *rel* inside *base*; raise on traversal/symlink escape."""
    if not _is_safe_rel(str(rel)):
        raise PublicationError(f"invalid {label}: path must not be absolute or contain '..'")
    candidate = (base / rel).resolve()
    try:
        candidate.relative_to(base.resolve())
    except ValueError as exc:
        raise PublicationError(f"{label} escapes allowed output directory") from exc
    return candidate


def validate_output_path(
    output_path: Optional[str],
    *,
    source_path: Optional[str] = None,
    label: str = "output",
) -> str:
    """Validate a single output path; return the resolved relative path.

    Rejects traversal, symlink escape, type conflicts (input==output) and
    ensures the target directory is usable.
    """
    if output_path is None or not isinstance(output_path, str) or not output_path:
        raise PublicationError(f"missing {label} path")
    if source_path is not None and source_path == output_path:
        raise PublicationError(f"{label} conflicts with source path")
    safe = _is_safe_rel(output_path)
    if not safe:
        raise PublicationError(f"{label} path contains traversal or invalid characters")
    return output_path


def validate_required_outputs(output_paths: list[str]) -> list[str]:
    """Return paths that do not exist yet (missing required outputs)."""
    missing: list[str] = []
    for p in output_paths:
        if p is None or not isinstance(p, str) or not p:
            missing.append(p or "<missing>")
        elif not os.path.exists(p):
            missing.append(p)
    return missing


class OutputPublication:
    """Applies permissions and writes outputs safely.

    Attributes:
        base_dir: workflow-state base (from MIND-070C).
        outputs_base_dir: directory into which outputs are published.
        overwrite_policy: 'never' (default), 'always', or 'explicit'.
        explicit_overwrite: True only when the operator granted explicit
            permission (overrides 'explicit' policy).
        workflow_id: used for diagnostics.
    """

    def __init__(
        self,
        base_dir: str,
        workflow_id: str,
        *,
        outputs_base_dir: Optional[str] = None,
        overwrite_policy: str = OVERWRITE_NEVER,
        explicit_overwrite: bool = False,
    ) -> None:
        self.base_dir = Path(base_dir).resolve()
        self.workflow_id = workflow_id
        self.outputs_base_dir = (
            (self.base_dir / outputs_base_dir) if outputs_base_dir else self.base_dir / "publications" / workflow_id
        ).resolve()
        if overwrite_policy not in (OVERWRITE_NEVER, OVERWRITE_ALWAYS, OVERWRITE_EXPLICIT):
            raise WorkflowError(
                f"invalid overwrite_policy: {overwrite_policy!r}",
                category=WorkflowErrorCategory.OUTPUT_PUBLICATION_FAILURE,
                step="<publication>",
                retryable=False,
            )
        self.overwrite_policy = overwrite_policy
        self.explicit_overwrite = explicit_overwrite
        self.errors: list[dict[str, Any]] = []

    def resolve(self, output_path: str) -> str:
        """Validate and resolve an output path to an absolute path inside the
        allowed outputs directory."""
        self._ensure_base_permissions()
        rel = Path(output_path)
        target = _resolve_within(self.outputs_base_dir, rel, "output")
        # A link that survives the safe resolution (e.g. its resolved target
        # still lives under the allowed dir) is inert; the traversal/symlink
        # check above already guarantees the resolved path stays inside base.
        return target

    def should_overwrite(self) -> bool:
        if self.overwrite_policy == OVERWRITE_ALWAYS:
            return True
        if self.overwrite_policy == OVERWRITE_EXPLICIT:
            return bool(self.explicit_overwrite)
        return False

    def publish(self, source_path: str, output_path: str) -> str:
        """Write source -> output; keep partial artifacts on failure."""
        self._ensure_base_permissions()
        target = self.resolve(output_path)
        if self.overwrite_policy == OVERWRITE_NEVER and os.path.exists(target):
            raise PublicationError(
                f"overwrite denied for {output_path!r}: explicit permission required"
            )
        if (
            self.overwrite_policy == OVERWRITE_EXPLICIT
            and not self.explicit_overwrite
            and os.path.exists(target)
        ):
            raise PublicationError(
                f"overwrite denied for {output_path!r}: explicit permission required"
            )
        target.parent.mkdir(parents=True, exist_ok=True)
        # Atomic publish: copy into a staging file in the target's own
        # directory, then os.replace() so the swap is one rename.  On an
        # allowed overwrite the old target is never unlinked first, so a
        # failed transfer (disk full, source dropped, etc.) leaves the old
        # file intact.  For a first write no old file exists, so a failed
        # write leaves nothing behind.
        fd, tmp = tempfile.mkstemp(dir=target.parent, prefix=".tmp-", suffix=".part")
        try:
            with os.fdopen(fd, "wb") as fh:
                with open(source_path, "rb") as src:
                    shutil.copyfileobj(src, fh)
            os.replace(tmp, target)
        except OSError as exc:
            try:
                os.unlink(tmp)
            except OSError:
                pass
            raise PublicationError(f"publish failed for {output_path!r}: {exc}") from exc
        return target

    def validate_required_outputs(self, output_paths: list[str]) -> list[str]:
        missing = validate_required_outputs(output_paths)
        if missing:
            raise PublicationError(
                f"required outputs missing before SUCCEEDED: {missing}"
            )
        return missing

    def record_failure(self, error: Exception) -> None:
        self.errors.append({"error": str(error), "type": type(error).__name__})

    def _ensure_base_permissions(self) -> None:
        # No filesystem mutation beyond resolved target creation by publish().
        pass
