from __future__ import annotations

from collections.abc import Callable

ProgressFn = Callable[[int, int, str], None]


def progress_bar(current: int, total: int, label: str = "", width: int = 40) -> None:
    filled = int(width * current / total) if total > 0 else width
    bar = "#" * filled + "-" * (width - filled)
    pct = f"{100 * current / total:.0f}%" if total > 0 else "100%"
    print(f"\r  [{bar}] {pct}  {label:<40}", end="", flush=True)
    if current >= total:
        print()


def noop_progress(current: int, total: int, label: str = "") -> None:
    del current, total, label
