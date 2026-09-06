"""Offline tests for the shared ProgressTracker emitter (runtime.operations)."""

from __future__ import annotations

import threading

from office365.runtime.operations import ProgressTracker


def test_report_emits_progress_with_total():
    seen = []
    tracker = ProgressTracker(seen.append, total=10, stage="uploading")

    tracker.report(3)
    tracker.report(5)

    assert [p.done for p in seen] == [3, 5]
    assert all(p.total == 10 for p in seen)  # noqa: PLR2004
    assert all(p.stage == "uploading" for p in seen)


def test_advance_and_late_total():
    seen = []
    tracker = ProgressTracker(seen.append, stage="migrating")

    tracker.advance()
    tracker.advance(2)
    tracker.set_total(10)
    tracker.report(5)

    assert [p.done for p in seen] == [1, 3, 5]
    assert seen[-1].total == 10  # noqa: PLR2004
    assert seen[-1].percent == 50.0  # noqa: PLR2004


def test_report_is_monotonic():
    seen = []
    tracker = ProgressTracker(seen.append, total=100, stage="loading")

    tracker.report(40)
    tracker.report(30)  # out-of-order completion must not go backwards

    assert [p.done for p in seen] == [40, 40]


def test_thread_safe_advances_sum_to_total():
    seen = []
    tracker = ProgressTracker(seen.append, total=1000, stage="parallel")
    threads = [threading.Thread(target=lambda: [tracker.advance() for _ in range(100)]) for _ in range(10)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert tracker.done == 1000  # noqa: PLR2004
    assert len(seen) == 1000  # noqa: PLR2004


def test_disabled_tracker_is_noop():
    tracker = ProgressTracker(None, total=5)
    tracker.report(1)
    assert tracker.done == 1
