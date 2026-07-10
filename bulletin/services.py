"""
DHIS2-like bulletin metrics engine.

Design choices (FDE rapid prototype):
- CSV stands in for DHIS2 API / Excel exports the MoH currently compiles manually.
- Pandas accelerates aggregation; in production this would move to a scheduled
  ETL job writing to HealthOS Data Models.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

DATA_FILE = Path(__file__).resolve().parent.parent / "data" / "dhis2_facility_indicators.csv"


@dataclass
class BulletinReport:
    current_period: str
    previous_period: str
    top_facilities: list[dict[str, Any]]
    maternal_summary: dict[str, Any]
    performance_scores: list[dict[str, Any]]
    trend_analysis: list[dict[str, Any]]
    generated_at: str


def _load_rows() -> list[dict[str, str]]:
    with DATA_FILE.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def _parse_date(value: str) -> datetime:
    return datetime.strptime(value, "%Y-%m-%d")


def _period_sort_key(period: str) -> tuple[int, int]:
    year, quarter = period[:4], period[-1]
    return int(year), int(quarter)


def _available_periods(rows: list[dict[str, str]]) -> list[str]:
    periods = sorted({row["period"] for row in rows}, key=_period_sort_key)
    return periods


def _rows_for_period(rows: list[dict[str, str]], period: str) -> list[dict[str, str]]:
    return [row for row in rows if row["period"] == period]


def _to_int(value: str) -> int:
    return int(value)


def top_facilities_by_volume(rows: list[dict[str, str]], limit: int = 10) -> list[dict[str, Any]]:
    ranked = sorted(rows, key=lambda r: _to_int(r["patient_volume"]), reverse=True)
    return [
        {
            "rank": index,
            "facility_id": row["facility_id"],
            "facility_name": row["facility_name"],
            "district": row["district"],
            "patient_volume": _to_int(row["patient_volume"]),
        }
        for index, row in enumerate(ranked[:limit], start=1)
    ]


def maternal_health_summary(rows: list[dict[str, str]]) -> dict[str, Any]:
    totals = {
        "anc_visits_first": sum(_to_int(r["anc_visits_first"]) for r in rows),
        "anc_visits_total": sum(_to_int(r["anc_visits_total"]) for r in rows),
        "deliveries": sum(_to_int(r["deliveries"]) for r in rows),
        "complications": sum(_to_int(r["complications"]) for r in rows),
    }
    deliveries = totals["deliveries"] or 1
    totals["complication_rate_pct"] = round((totals["complications"] / deliveries) * 100, 2)
    totals["facility_count"] = len(rows)
    return totals


def facility_performance_scores(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    scores: list[dict[str, Any]] = []
    for row in rows:
        submitted = _parse_date(row["report_submitted_date"])
        deadline = _parse_date(row["reporting_deadline"])
        on_time = submitted <= deadline
        # Completeness proxy: all mandatory numeric fields present and > 0
        completeness_fields = [
            row["patient_volume"],
            row["anc_visits_total"],
            row["deliveries"],
        ]
        complete = all(_to_int(value) > 0 for value in completeness_fields)
        timeliness_score = 100 if on_time else 40
        completeness_score = 100 if complete else 55
        overall = round((timeliness_score * 0.4) + (completeness_score * 0.6), 1)
        scores.append(
            {
                "facility_id": row["facility_id"],
                "facility_name": row["facility_name"],
                "district": row["district"],
                "timeliness_score": timeliness_score,
                "completeness_score": completeness_score,
                "overall_score": overall,
                "submitted_on_time": on_time,
                "report_complete": complete,
            }
        )
    return sorted(scores, key=lambda item: item["overall_score"], reverse=True)


def quarter_over_quarter_trends(
    current_rows: list[dict[str, str]],
    previous_rows: list[dict[str, str]],
) -> list[dict[str, Any]]:
    previous_by_facility = {row["facility_id"]: row for row in previous_rows}
    trends: list[dict[str, Any]] = []

    for row in current_rows:
        facility_id = row["facility_id"]
        prev = previous_by_facility.get(facility_id)
        if not prev:
            continue
        current_volume = _to_int(row["patient_volume"])
        previous_volume = _to_int(prev["patient_volume"])
        delta = current_volume - previous_volume
        pct_change = round((delta / previous_volume) * 100, 2) if previous_volume else 0.0
        trends.append(
            {
                "facility_id": facility_id,
                "facility_name": row["facility_name"],
                "district": row["district"],
                "current_volume": current_volume,
                "previous_volume": previous_volume,
                "volume_change": delta,
                "volume_change_pct": pct_change,
                "deliveries_change": _to_int(row["deliveries"]) - _to_int(prev["deliveries"]),
            }
        )
    return sorted(trends, key=lambda item: abs(item["volume_change_pct"]), reverse=True)[:10]


def build_bulletin(current_period: str | None = None) -> BulletinReport:
    rows = _load_rows()
    periods = _available_periods(rows)
    if not periods:
        raise ValueError("No reporting periods found in sample data.")

    current = current_period or periods[-1]
    if current not in periods:
        raise ValueError(f"Unknown period '{current}'. Available: {', '.join(periods)}")

    current_index = periods.index(current)
    previous = periods[current_index - 1] if current_index > 0 else current

    current_rows = _rows_for_period(rows, current)
    previous_rows = _rows_for_period(rows, previous)

    return BulletinReport(
        current_period=current,
        previous_period=previous,
        top_facilities=top_facilities_by_volume(current_rows),
        maternal_summary=maternal_health_summary(current_rows),
        performance_scores=facility_performance_scores(current_rows),
        trend_analysis=quarter_over_quarter_trends(current_rows, previous_rows),
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )
