"""
PACE Payroll Validator — Coverage Validator
=============================================
Determines whether payroll periods fully cover a policy period.

Statuses:
    FULL    — Policy period is covered, with edge gaps ≤ BUFFER_DAYS tolerated.
    PARTIAL — Gaps exist beyond buffer tolerance (internal OR edge gaps too large).
    NO      — Zero overlap between payroll and policy periods.

Buffer Logic:
    After finding all gaps within [policy_start, policy_end]:
      - A gap at the VERY START (gap_start == policy_start) that is ≤ BUFFER_DAYS wide → forgiven.
      - A gap at the VERY END   (gap_end   == policy_end)   that is ≤ BUFFER_DAYS wide → forgiven.
      - Internal gaps between payroll segments are NEVER forgiven regardless of size.
    If no unforgiven gaps remain → FULL. Otherwise → PARTIAL.
"""

from datetime import date, timedelta
from typing import List, Tuple

from config import BUFFER_DAYS
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
        (status, remaining_gaps)
        - status         : "FULL", "PARTIAL", or "NO"
        - remaining_gaps : gaps that were NOT forgiven by the buffer
    """
    if policy_start > policy_end:
        logger.error(
            "Invalid policy period: policy_start (%s) is after policy_end (%s)",
            policy_start, policy_end,
        )
        return ("NO", [(policy_start, policy_end)])

    # Filter out malformed payroll periods (start > end)
    cleaned_periods = []
    for start, end in payroll_periods:
        if start <= end:
            cleaned_periods.append((start, end))
        else:
            logger.warning(
                "Ignored invalid payroll period: %s -> %s (start is after end)",
                start, end,
            )

    if not cleaned_periods:
        return ("NO", [(policy_start, policy_end)])

    # 1. Merge overlapping / adjacent payroll periods
    merged = _merge_periods(cleaned_periods)

    # 2. Clip merged periods to the policy window [policy_start, policy_end]
    clipped = _clip_to_policy(merged, policy_start, policy_end)

    if not clipped:
        # No overlap at all
        return ("NO", [(policy_start, policy_end)])

    # 3. Find raw gaps within the policy window
    raw_gaps = _find_gaps(clipped, policy_start, policy_end)

    if not raw_gaps:
        return ("FULL", [])

    # 4. Apply buffer: forgive start/end edge gaps that are within BUFFER_DAYS
    remaining_gaps = []
    for g_start, g_end in raw_gaps:
        gap_days = (g_end - g_start).days + 1  # inclusive day count
        is_start_edge = (g_start == policy_start)
        is_end_edge   = (g_end   == policy_end)

        if (is_start_edge or is_end_edge) and gap_days <= BUFFER_DAYS:
            logger.info(
                "Buffer absorbed %s edge gap: %s -> %s (%d day(s), tolerance=%d)",
                "start" if is_start_edge else "end",
                g_start, g_end, gap_days, BUFFER_DAYS,
            )
        else:
            remaining_gaps.append((g_start, g_end))

    if not remaining_gaps:
        return ("FULL", [])
    else:
        return ("PARTIAL", remaining_gaps)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _merge_periods(
    periods: List[Tuple[date, date]],
) -> List[Tuple[date, date]]:
    """
    Sort and merge overlapping or adjacent date ranges.
    Two periods are adjacent if one ends on day D and the next starts on D or D+1.
    """
    sorted_p = sorted(periods, key=lambda t: t[0])
    merged: List[Tuple[date, date]] = [sorted_p[0]]

    for start, end in sorted_p[1:]:
        prev_start, prev_end = merged[-1]
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
    """Clip merged periods to the policy window, dropping non-overlapping ranges."""
    clipped: List[Tuple[date, date]] = []
    for start, end in merged:
        if end < policy_start or start > policy_end:
            continue
        clipped.append((max(start, policy_start), min(end, policy_end)))
    return clipped


def _find_gaps(
    clipped: List[Tuple[date, date]],
    policy_start: date,
    policy_end: date,
) -> List[Tuple[date, date]]:
    """Identify all date gaps within [policy_start, policy_end]."""
    gaps: List[Tuple[date, date]] = []

    # Gap before the first clipped period
    if clipped[0][0] > policy_start:
        gaps.append((policy_start, clipped[0][0] - timedelta(days=1)))

    # Internal gaps between consecutive clipped periods
    for i in range(len(clipped) - 1):
        gap_start = clipped[i][1] + timedelta(days=1)
        gap_end   = clipped[i + 1][0] - timedelta(days=1)
        if gap_start <= gap_end:
            gaps.append((gap_start, gap_end))

    # Gap after the last clipped period
    if clipped[-1][1] < policy_end:
        gaps.append((clipped[-1][1] + timedelta(days=1), policy_end))

    return gaps
