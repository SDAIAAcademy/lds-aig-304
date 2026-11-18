# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""Sampling numbers quoted on the day 3 slides and in lab 3.2 (exact, no simulation needed).

Run from the course root:  uv run labs/_tools/sampling.py
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import case_data as D  # noqa: E402


def zero_defect_n(rate, confidence=0.95):
    """Smallest sample in which at least one exception is found with the given confidence,
    when the population has the given exception rate (large population)."""
    return math.ceil(math.log(1 - confidence) / math.log(1 - rate))


def p_at_least_one(population, bad, n):
    """Hypergeometric: chance that a random sample of n rows holds at least one bad row."""
    return 1 - math.comb(population - bad, n) / math.comb(population, n)


def working_days(start, end):
    """Working days after start up to and including end; the week runs Sunday to Thursday."""
    import datetime as dt
    d0, d1 = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    n, d = 0, d0
    while d < d1:
        d += dt.timedelta(days=1)
        if d.weekday() not in (4, 5):  # Friday, Saturday
            n += 1
    return n


def main():
    for rate in (0.20, 0.10, 0.05, 0.02):  # 20% is practice question q-d3-05
        print(f"zero-defect sample for 95% confidence at {rate:.0%}: {zero_defect_n(rate)}")
    pop, bad = D.OVERRIDES_TOTAL, D.OVERRIDES_BY_UNTRAINED
    for n in (5, 10, 20, 29):
        print(f"overrides: {bad} of {pop} by untrained reviewers; sample {n}: P(at least one) = {p_at_least_one(pop, bad, n):.1%}")
    # parallel example on the Day 3 slides (equipment-rental company): 15 faulty of 300 contracts
    for n in (5, 10, 20, 29):
        print(f"slides example: 15 of 300; sample {n}: P(at least one) = {p_at_least_one(300, 15, n):.1%}")
    inc, rec, got = D.INCIDENT_LATE
    print(f"{inc}: {working_days(rec, got)} working days")
    days = [working_days(r, g) for _, r, g in D.FOLLOWUP_INCIDENTS]
    ok = sum(d <= D.INCIDENT_LIMIT_WORKING_DAYS for d in days)
    print(f"follow-up: {ok} of {len(days)} within {D.INCIDENT_LIMIT_WORKING_DAYS} working days = {ok / len(days):.1%}; days {days}")


if __name__ == "__main__":
    main()
