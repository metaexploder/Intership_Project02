"""
PACE Payroll Validator — Coverage Validator
=============================================
Determines whether payroll periods fully cover a policy period.

Statuses:
    FULL    — All three conditions met:
                1. abs(InceptionDate - PeriodStart) <= BUFFER_DAYS
                2. abs(PeriodEnd - ExpirationDate)  <= BUFFER_DAYS
                3. No internal gaps within [InceptionDate, ExpirationDate]
    PARTIAL — At least one condition above is violated.
    NO      — Zero overlap between payroll and policy periods.

Buffer Definitions:
    StartBuffer = InceptionDate - PeriodStart
        Positive → payroll starts BEFORE inception (early coverage, surplus)
        Negative → payroll starts AFTER  inception (late start, deficit)

    EndBuffer = PeriodEnd - ExpirationDate
        Positive → payroll ends AFTER  expiration (extra coverage, surplus)
        Negative → payroll ends BEFORE expiration (ends early, deficit)

    FULL requires: abs(StartBuffer) <= BUFFER_DAYS AND abs(EndBuffer) <= BUFFER_DAYS
                   AND no internal gaps.
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
    Validate payroll coverage against the policy period with ± BUFFER_DAYS tolerance.

    Parameters
    ----------
    policy_start : date   — InceptionDate from DB
    policy_end   : date   — ExpirationDate from DB
    payroll_periods : List[Tuple[date, date]]
        Each tuple is (period_start, period_end) from the payroll Excel files.

    Returns
    -------
    Tuple[str, List[Tuple[date, date]]]
        (status, unforgiven_gaps)
        status          : "FULL", "PARTIAL", or "NO"
        unforgiven_gaps : gaps that are NOT excused by the buffer
    """
    # --- Guard: invalid policy window ---
    if policy_start > policy_end:
        logger.error(
            "Invalid policy period: policy_start (%s) is after policy_end (%s)",
            policy_start, policy_end,
        )
        return ("NO", [(policy_start, policy_end)])

    # --- Filter out malformed payroll periods (start > end) ---
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

    # -----------------------------------------------------------------------
    # Step 1 — Compute overall payroll bounds (PeriodStart / PeriodEnd)
    # -----------------------------------------------------------------------
    period_start: date = min(s for s, _ in cleaned_periods)
    period_end:   date = max(e for _, e in cleaned_periods)

    # -----------------------------------------------------------------------
    # Step 2 — Compute buffers
    # -----------------------------------------------------------------------
    start_buffer: int = (policy_start - period_start).days  # +ve = payroll starts before inception
    end_buffer:   int = (period_end   - policy_end).days    # +ve = payroll ends after expiration

    start_ok: bool = abs(start_buffer) <= BUFFER_DAYS
    end_ok:   bool = abs(end_buffer)   <= BUFFER_DAYS

    logger.debug(
        "WOID buffers: PeriodStart=%s PeriodEnd=%s | StartBuffer=%+dd EndBuffer=%+dd | "
        "StartOK=%s EndOK=%s",
        period_start, period_end, start_buffer, end_buffer, start_ok, end_ok,
    )

    if not start_ok:
        logger.info(
            "StartBuffer %+d days exceeds ±%d day tolerance "
            "(InceptionDate=%s, PeriodStart=%s)",
            start_buffer, BUFFER_DAYS, policy_start, period_start,
        )
    if not end_ok:
        logger.info(
            "EndBuffer %+d days exceeds ±%d day tolerance "
            "(ExpirationDate=%s, PeriodEnd=%s)",
            end_buffer, BUFFER_DAYS, policy_end, period_end,
        )

    # -----------------------------------------------------------------------
    # Step 3 — Merge payroll periods & clip to policy window
    # -----------------------------------------------------------------------
    merged  = _merge_periods(cleaned_periods)
    clipped = _clip_to_policy(merged, policy_start, policy_end)

    if not clipped:
        # Zero overlap even after all processing → NO
        return ("NO", [(policy_start, policy_end)])

    # -----------------------------------------------------------------------
    # Step 4 — Find all gaps within [policy_start, policy_end]
    # -----------------------------------------------------------------------
    all_gaps = _find_gaps(clipped, policy_start, policy_end)

    # Separate into internal gaps and edge gaps
    # Internal gaps: start AFTER policy_start AND end BEFORE policy_end
    internal_gaps = [
        (gs, ge) for gs, ge in all_gaps
        if gs > policy_start and ge < policy_end
    ]

    if internal_gaps:
        for gs, ge in internal_gaps:
            logger.info("Internal gap found: %s -> %s", gs, ge)

    # -----------------------------------------------------------------------
    # Step 5 — Determine status
    # -----------------------------------------------------------------------
    if start_ok and end_ok and not internal_gaps:
        return ("FULL", [])

    # Build the list of unforgiven gaps for reporting
    unforgiven: List[Tuple[date, date]] = []
    for gs, ge in all_gaps:
        is_start_edge = (gs == policy_start)
        is_end_edge   = (ge == policy_end)

        if is_start_edge and start_ok:
            # Edge gap at start is within buffer tolerance — forgiven
            logger.debug("Start edge gap %s->%s forgiven by buffer.", gs, ge)
            continue
        if is_end_edge and end_ok:
            # Edge gap at end is within buffer tolerance — forgiven
            logger.debug("End edge gap %s->%s forgiven by buffer.", gs, ge)
            continue
        unforgiven.append((gs, ge))

    return ("PARTIAL", unforgiven)


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
    """Clip merged periods to [policy_start, policy_end], dropping non-overlapping ranges."""
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
    """Find all date gaps within [policy_start, policy_end]."""
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
