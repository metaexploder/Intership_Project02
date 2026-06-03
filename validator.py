"""
PACE Payroll Validator — Coverage Validator
=============================================
Determines whether payroll periods fully cover a policy period.

Statuses:
    FULL    — Every day of the policy period is covered.
    PARTIAL — Some policy dates are covered; gaps exist.
    NO      — Zero overlap between payroll and policy periods.
"""

from datetime import date, timedelta
from typing import List, Tuple

from logger_setup import setup_logger

logger = setup_logger()


def validate_coverage(
    policy_start: date,
    policy_end: date,
    payroll_periods: List[Tuple[date, date]],
) -> Tuple[str, List[Tuple[date, date]]]:
    """
    Validate payroll coverage against the policy period.

    Parameters
    ----------
    policy_start : date
    policy_end : date
    payroll_periods : List[Tuple[date, date]]
        Each tuple is (period_start, period_end).

    Returns
    -------
    Tuple[str, List[Tuple[date, date]]]
        (status, missing_gaps)
        - status: "FULL", "PARTIAL", or "NO"
        - missing_gaps: list of (gap_start, gap_end) tuples
    """
    if not payroll_periods:
        gap = (policy_start, policy_end)
        return ("NO", [gap])

    # 1. Merge overlapping / adjacent payroll periods
    merged = _merge_periods(payroll_periods)

    # 2. Clip merged periods to the policy window
    clipped = _clip_to_policy(merged, policy_start, policy_end)

    if not clipped:
        # No overlap at all
        gap = (policy_start, policy_end)
        return ("NO", [gap])

    # 3. Find gaps within the policy window
    gaps = _find_gaps(clipped, policy_start, policy_end)

    if not gaps:
        return ("FULL", [])
    else:
        # Check if clipped periods provide *any* coverage
        # (they do, because clipped is non-empty)
        return ("PARTIAL", gaps)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _merge_periods(
    periods: List[Tuple[date, date]],
) -> List[Tuple[date, date]]:
    """
    Sort and merge overlapping or adjacent date ranges.

    Two periods are considered adjacent if one ends on day D and the
    next starts on day D or D+1.
    """
    sorted_p = sorted(periods, key=lambda t: t[0])
    merged: List[Tuple[date, date]] = [sorted_p[0]]

    for start, end in sorted_p[1:]:
        prev_start, prev_end = merged[-1]
        # Adjacent: prev_end + 1 day >= start  → merge
        if start <= prev_end + timedelta(days=1):
            merged[-1] = (prev_start, max(prev_end, end))
        else:
            merged.append((start, end))

    return merged


def _clip_to_policy(
    merged: List[Tuple[date, date]],
    policy_start: date,
    policy_end: date,
) -> List[Tuple[date, date]]:
    """Clip merged periods to the policy window, dropping non-overlapping."""
    clipped: List[Tuple[date, date]] = []
    for start, end in merged:
        # No overlap
        if end < policy_start or start > policy_end:
            continue
        clipped_start = max(start, policy_start)
        clipped_end = min(end, policy_end)
        clipped.append((clipped_start, clipped_end))
    return clipped


def _find_gaps(
    clipped: List[Tuple[date, date]],
    policy_start: date,
    policy_end: date,
) -> List[Tuple[date, date]]:
    """
    Identify date gaps between the clipped periods within the
    policy window.
    """
    gaps: List[Tuple[date, date]] = []

    # Gap before the first clipped period
    if clipped[0][0] > policy_start:
        gaps.append((policy_start, clipped[0][0] - timedelta(days=1)))

    # Gaps between consecutive clipped periods
    for i in range(len(clipped) - 1):
        gap_start = clipped[i][1] + timedelta(days=1)
        gap_end = clipped[i + 1][0] - timedelta(days=1)
        if gap_start <= gap_end:
            gaps.append((gap_start, gap_end))

    # Gap after the last clipped period
    if clipped[-1][1] < policy_end:
        gaps.append((clipped[-1][1] + timedelta(days=1), policy_end))

    return gaps
