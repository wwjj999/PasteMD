"""macOS keystroke helpers (via AppleScript/System Events)."""

from __future__ import annotations

import subprocess

from ...core.errors import ClipboardError


def simulate_paste(*, timeout_s: float = 5.0) -> None:
    """
    Simulate ⌘V in the frontmost app.

    Requires: System Settings → Privacy & Security → Accessibility (and possibly Automation).
    """
    script = 'tell application "System Events" to keystroke "v" using command down'
    try:
        subprocess.run(
            ["osascript", "-e", script],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_s,
        )
    except subprocess.CalledProcessError as e:
        msg = (e.stderr or "").strip() or (e.stdout or "").strip() or str(e)
        raise ClipboardError(f"Failed to simulate Cmd+V: {msg}") from e
    except subprocess.TimeoutExpired as e:
        raise ClipboardError("Failed to simulate Cmd+V: timeout") from e
