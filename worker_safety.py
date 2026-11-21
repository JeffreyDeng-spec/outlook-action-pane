"""Portable process-level safety primitives for Outlook Triage.

These helpers do not connect to Microsoft or load real account credentials.
"""
from __future__ import annotations

import contextlib
import os
from pathlib import Path

class LockBusy(RuntimeError):
    """Another process already owns the lock."""

@contextlib.contextmanager
def exclusive_lock(path: Path):
    """Hold a whole-worker lock across its startup, requests and background jobs."""
    path = Path(path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as handle:
        # A Windows byte-range lock may cover a byte beyond EOF. Do not
        # pre-write that byte: concurrent first openers could race an owner.
        handle.seek(0)
        if os.name == "nt":
            import msvcrt
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc:
                if exc.errno not in (13, 36):
                    raise
                raise LockBusy(str(path)) from exc
            try:
                yield
            finally:
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError as exc:
                raise LockBusy(str(path)) from exc
            try:
                yield
            finally:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)

def worker_state_root(value: str | None = None) -> Path:
    """Workers use an explicit absolute state directory, never a legacy default."""
    value = value or os.environ.get("OUTLOOK_TRIAGE_STATE_DIR")
    if not value or not Path(value).is_absolute():
        raise ValueError("Worker requires an absolute OUTLOOK_TRIAGE_STATE_DIR")
    return Path(value).resolve()

def account_state(root: Path, account: str, directory_key: str) -> Path:
    """Keep explicit configured account folders under the worker state root."""
    if not account or not directory_key or any(c in directory_key for c in "/\\:\0"):
        raise ValueError("Unsafe account state directory")
    if directory_key in (".", "..") or directory_key != directory_key.strip(" ."):
        raise ValueError("Unsafe account state directory")
    root = Path(root).resolve()
    target = (root / directory_key).resolve()
    if target.parent != root or target == root:
        raise ValueError("Account path is not beneath the worker state root")
    return target
