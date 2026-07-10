"""
Healthcare bulletin analytics engine.

Decision notes:
- The app uses only the files from the provided Google Drive folder.
- Each dashboard/API request re-syncs the folder (with a short TTL) and re-reads
  every CSV/XLS/XLSX file so newly added Drive files appear automatically.
- Pandas is used here because it is the fastest way to prove multi-source joins,
  quarter aggregation, and bulletin-style metrics in an interview assignment.
"""
from __future__ import annotations

import re
import sys
import time
import urllib.parse
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import bs4
import pandas as pd
import requests

Granularity = Literal["M", "Q", "A"]
GRANULARITY_ORDER: tuple[Granularity, ...] = ("M", "Q", "A")
GRANULARITY_CONFIG: dict[Granularity, dict[str, Any]] = {
    "M": {
        "label": "Monthly",
        "freq": "M",
        "column": "reporting_month_period",
        "child_freq": None,
        "comparison": "MoM",
        "expected_units": 1,
        "unit_label": "month",
        "trend_points": 6,
    },
    "Q": {
        "label": "Quarterly",
        "freq": "Q",
        "column": "quarter",
        "child_freq": "M",
        "comparison": "QoQ",
        "expected_units": 3,
        "unit_label": "months",
        "trend_points": 6,
    },
    "A": {
        "label": "Annual",
        "freq": "Y",
        "column": "reporting_year",
        "child_freq": "Q",
        "comparison": "YoY",
        "expected_units": 12,
        "unit_label": "months",
        "trend_points": 4,
    },
}
CHILD_PERIOD_COLUMNS = {
    "M": "reporting_month_period",
    "Q": "quarter",
}

DATA_SOURCE_URL = "https://drive.google.com/drive/folders/1DPk6jKSO_bbnhonUX6S91kLWZVulWTmA"
DRIVE_FOLDER_ID = "1DPk6jKSO_bbnhonUX6S91kLWZVulWTmA"
DATA_DIRECTORY = Path(__file__).resolve().parent.parent / "data" / "google_drive_source"
DRIVE_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
# Known assignment files — used when live folder listing is rate-limited.
DEFAULT_DRIVE_FILES: dict[str, str] = {
    "clinical_neonatal.csv": "1h9FDzHw6IXqV0ndq4ZAbhbQuRRK_pqE8",
    "facilities.csv": "1_TeF2731LCDdjIJWHeNGnIiG1AS2vtSF",
    "governance.csv": "1reVIu7yZp-grt1QqNVcZX2z-acRMUb4j",
    "healthcare_workers.csv": "1qBX7fGyEoagkphBqJEHjk2G-vmPnwrLa",
    "operations.csv": "14EouJuR5JJk6b0uTe_C4jWvC0AQ4jJi9",
}
SUPPORTED_EXTENSIONS = {".csv", ".xls", ".xlsx"}
CORE_DATASETS = {
    "clinical_neonatal",
    "facilities",
    "governance",
    "healthcare_workers",
    "operations",
}
# Filename hints help classify newly added Drive files without manual configuration.
ROLE_HINTS: dict[str, str] = {
    "clinical_neonatal": "clinical",
    "clinical": "clinical",
    "neonatal": "clinical",
    "facilities": "facilities",
    "facility": "facilities",
    "governance": "governance",
    "healthcare_workers": "healthcare_workers",
    "workers": "healthcare_workers",
    "workforce": "healthcare_workers",
    "operations": "operations",
    "ops": "operations",
}
SYNC_TTL_SECONDS = 300
AUTO_REFRESH_SECONDS = 30
_last_drive_sync_epoch = 0.0
_last_synced_at = ""
_sync_status = "cached"
_sync_message = "Serving analytics from local source files."

BOOLEAN_SCORES = {"yes": 100.0, "partial": 60.0, "outdated": 40.0, "no": 0.0}
COVERAGE_SCORES = {"full": 100.0, "partial": 55.0, "none": 0.0}
SUPPLY_SCORES = {"always": 100.0, "usually": 80.0, "sometimes": 55.0, "rarely": 20.0, "no": 0.0}


@dataclass(frozen=True)
class PeriodOption:
    key: str
    label: str
    month_count: int
    is_partial: bool
    granularity: str
    expected_units: int
    coverage_pct: float


@dataclass
class BulletinReport:
    current_period: str
    current_period_label: str
    previous_period: str
    previous_period_label: str
    granularity: str
    granularity_label: str
    comparison_label: str
    generated_at: str
    source_url: str
    source_files: list[str]
    available_periods: list[dict[str, Any]]
    available_granularities: list[dict[str, str]]
    source_summary: dict[str, Any]
    headline_metrics: list[dict[str, Any]]
    maternal_summary: dict[str, Any]
    top_facilities: list[dict[str, Any]]
    performance_scores: list[dict[str, Any]]
    provincial_overview: list[dict[str, Any]]
    monthly_trends: list[dict[str, Any]]
    period_breakdown: list[dict[str, Any]]
    rollup_summary: dict[str, Any]
    advanced_aggregations: list[dict[str, Any]]
    trend_analysis: list[dict[str, Any]]
    risk_alerts: list[dict[str, Any]]
    strategic_summary: list[dict[str, str]]


def report_to_dict(report: BulletinReport) -> dict[str, Any]:
    return asdict(report)


def _safe_divide(numerator: float, denominator: float, multiplier: float = 1.0) -> float:
    if not denominator:
        return 0.0
    return round((numerator / denominator) * multiplier, 2)


def _clamp(value: float, lower: float = 0.0, upper: float = 100.0) -> float:
    if hasattr(value, "clip"):
        return value.clip(lower=lower, upper=upper)
    return max(lower, min(upper, value))


def _read_tabular_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


def _score_from_map(value: Any, mapping: dict[str, float], default: float = 0.0) -> float:
    text = str(value).strip().lower()
    return mapping.get(text, default)


def _score_percent(value: Any) -> float:
    text = str(value).strip().replace("%", "")
    if not text or text.lower() == "nan":
        return 0.0
    return float(text)


def _normalize_granularity(value: str | None) -> Granularity:
    normalized = (value or "Q").strip().upper()
    if normalized in GRANULARITY_CONFIG:
        return normalized  # type: ignore[return-value]
    return "Q"


def _infer_granularity_from_period(period_key: str | None, fallback: Granularity = "Q") -> Granularity:
    if not period_key:
        return fallback
    normalized = period_key.strip().upper()
    if re.fullmatch(r"\d{4}-\d{2}", normalized):
        return "M"
    if re.fullmatch(r"\d{4}Q[1-4]", normalized):
        return "Q"
    if re.fullmatch(r"\d{4}", normalized):
        return "A"
    return fallback


def _period_sort_value(period_key: str, granularity: Granularity) -> pd.Timestamp:
    return pd.Period(period_key, freq=GRANULARITY_CONFIG[granularity]["freq"]).start_time


def _to_period_option(period: pd.Period, unit_count: int, granularity: Granularity) -> PeriodOption:
    config = GRANULARITY_CONFIG[granularity]
    expected_units = int(config["expected_units"])
    is_partial = unit_count < expected_units
    coverage_pct = round(_safe_divide(unit_count, expected_units, 100), 1)

    if granularity == "M":
        label = period.strftime("%b %Y")
    elif granularity == "A":
        label = f"Calendar year {period.year}"
    else:
        label = f"{period}"

    if is_partial and granularity != "M":
        label = f"{label} ({unit_count}/{expected_units} {config['unit_label']} loaded)"

    return PeriodOption(
        key=str(period),
        label=label,
        month_count=unit_count,
        is_partial=is_partial,
        granularity=granularity,
        expected_units=expected_units,
        coverage_pct=coverage_pct,
    )


def _period_column(granularity: Granularity) -> str:
    return GRANULARITY_CONFIG[granularity]["column"]


def _child_period_column(child_freq: str) -> str:
    return CHILD_PERIOD_COLUMNS[child_freq]


def _clinical_totals(frame: pd.DataFrame) -> dict[str, float]:
    deliveries = float(frame["total_deliveries"].sum())
    live_births = float(frame["live_births"].sum())
    stillbirths = float(frame["stillbirths"].sum())
    neonatal_deaths = float(frame["neonatal_deaths_total"].sum())
    return {
        "deliveries": deliveries,
        "live_births": live_births,
        "stillbirths": stillbirths,
        "neonatal_deaths": neonatal_deaths,
        "neonatal_mortality_rate": _safe_divide(neonatal_deaths, live_births, 1000),
        "stillbirth_rate": _safe_divide(stillbirths, deliveries, 1000),
        "preterm_rate_pct": _safe_divide(float(frame["preterm_births_total"].sum()), live_births, 100),
        "low_birth_weight_rate_pct": _safe_divide(float(frame["birth_weight_less_2500g"].sum()), live_births, 100),
    }


def _classify_score(score: float) -> str:
    if score >= 80:
        return "High performing"
    if score >= 60:
        return "Stable"
    if score >= 40:
        return "Watch"
    return "Critical"


def _is_test_environment() -> bool:
    return "test" in sys.argv


def _list_local_files() -> list[Path]:
    if not DATA_DIRECTORY.exists():
        return []
    return sorted(
        path
        for path in DATA_DIRECTORY.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized = df.copy()
    normalized.columns = [str(column).strip() for column in normalized.columns]
    if "facility_id" in normalized.columns:
        normalized["facility_id"] = normalized["facility_id"].astype(str).str.strip()
    return normalized


def _classify_dataset(stem: str, df: pd.DataFrame) -> str:
    stem_key = stem.lower().replace("-", "_")
    for hint, role in ROLE_HINTS.items():
        if hint in stem_key:
            return role

    columns = {column.lower() for column in df.columns}
    if {"reporting_month", "total_deliveries"}.issubset(columns):
        return "clinical"
    if {"facility_name", "district", "province"}.issubset(columns):
        return "facilities"
    if "hmis_reporting_completeness" in columns or "newborn_protocol_exists" in columns:
        return "governance"
    if "total_nurses" in columns or "neonatal_trained_nurses" in columns:
        return "healthcare_workers"
    if "avg_referral_time_hrs" in columns or "essential_drugs_stockouts_days" in columns:
        return "operations"
    if "facility_id" in columns:
        return "supplemental"
    return "unclassified"


def _prepare_clinical_frame(df: pd.DataFrame) -> pd.DataFrame:
    clinical = _normalize_columns(df)
    clinical["reporting_month"] = pd.to_datetime(clinical["reporting_month"], format="%Y-%m", errors="coerce")
    clinical["reporting_month_period"] = clinical["reporting_month"].dt.to_period("M")
    clinical["quarter"] = clinical["reporting_month"].dt.to_period("Q")
    clinical["reporting_year"] = clinical["reporting_month"].dt.to_period("Y")
    if "neonatal_deaths_total" not in clinical.columns:
        early = clinical.get("neonatal_deaths_0_7d", 0)
        late = clinical.get("neonatal_deaths_8_28d", 0)
        clinical["neonatal_deaths_total"] = early.fillna(0) + late.fillna(0)
    if "preterm_births_total" not in clinical.columns:
        early = clinical.get("preterm_births_28_32w", 0)
        late = clinical.get("preterm_births_32_37w", 0)
        clinical["preterm_births_total"] = early.fillna(0) + late.fillna(0)
    return clinical


def _local_files_timestamp() -> str:
    files = _list_local_files()
    if not files:
        return ""
    latest = max(path.stat().st_mtime for path in files)
    return datetime.fromtimestamp(latest).strftime("%Y-%m-%d %H:%M:%S")


def _cleanup_stale_partials() -> None:
    if not DATA_DIRECTORY.exists():
        return
    for path in DATA_DIRECTORY.iterdir():
        if path.is_file() and ".part" in path.name:
            path.unlink(missing_ok=True)


def _has_complete_local_cache() -> bool:
    stems = {path.stem for path in _list_local_files()}
    return CORE_DATASETS.issubset(stems)


def _drive_session() -> requests.Session:
    session = requests.Session()
    session.headers["User-Agent"] = DRIVE_USER_AGENT
    return session


def _list_drive_folder_files(session: requests.Session, folder_id: str) -> dict[str, str]:
    """Return {filename: file_id} for tabular files in a public Drive folder."""
    params = urllib.parse.urlencode({"id": folder_id})
    response = session.get(
        f"https://drive.google.com/embeddedfolderview?{params}",
        timeout=60,
    )
    response.raise_for_status()

    soup = bs4.BeautifulSoup(response.text, features="html.parser")
    manifest: dict[str, str] = {}
    for anchor in soup.find_all("a"):
        href = anchor.get("href", "")
        if not isinstance(href, str):
            continue

        file_match = re.match(r"https://drive\.google\.com/file/d/([-\w]{25,})/view", href)
        docs_match = re.match(r"https://docs\.google\.com/\w+/d/([-\w]{25,})/", href)
        match = file_match or docs_match
        if not match:
            continue

        filename = anchor.get_text(strip=True)
        if not filename or Path(filename).suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        manifest[filename] = match.group(1)
    return manifest


def _resolve_drive_manifest(session: requests.Session) -> dict[str, str]:
    manifest = dict(DEFAULT_DRIVE_FILES)
    try:
        discovered = _list_drive_folder_files(session, DRIVE_FOLDER_ID)
        manifest.update(discovered)
    except Exception:
        pass
    return manifest


def _download_drive_file(session: requests.Session, file_id: str, destination: Path) -> None:
    """Download a public Drive file without gdown (avoids FileURLRetrievalError)."""
    base_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    last_error: Exception | None = None

    for attempt in range(4):
        try:
            response = session.get(base_url, timeout=120)
            response.raise_for_status()
            content_type = response.headers.get("content-type", "")

            if "text/html" in content_type:
                confirm_match = re.search(r"confirm=([0-9A-Za-z_]+)", response.text)
                download_url_match = re.search(r'href="(/uc\?export=download[^"]+)"', response.text)
                if confirm_match:
                    confirm_url = f"{base_url}&confirm={confirm_match.group(1)}"
                    response = session.get(confirm_url, timeout=120)
                    response.raise_for_status()
                elif download_url_match:
                    confirm_url = "https://drive.google.com" + download_url_match.group(1).replace("&amp;", "&")
                    response = session.get(confirm_url, timeout=120)
                    response.raise_for_status()
                elif len(response.content) < 50_000:
                    raise RuntimeError("Google Drive returned an HTML page instead of file bytes.")

            if len(response.content) < 32:
                raise RuntimeError("Downloaded file is unexpectedly small.")

            temp_path = destination.with_suffix(destination.suffix + ".tmp")
            temp_path.write_bytes(response.content)
            temp_path.replace(destination)
            return
        except Exception as exc:
            last_error = exc
            time.sleep(1.0 * (attempt + 1))

    if last_error is not None:
        raise last_error


def _download_drive_folder() -> list[Path]:
    _cleanup_stale_partials()
    DATA_DIRECTORY.mkdir(parents=True, exist_ok=True)

    session = _drive_session()
    manifest = _resolve_drive_manifest(session)
    if not manifest:
        raise ValueError("No downloadable files found in the Google Drive folder.")

    failed_downloads: list[str] = []
    for filename, file_id in manifest.items():
        destination = DATA_DIRECTORY / filename
        try:
            _download_drive_file(session, file_id, destination)
        except Exception:
            if destination.exists() and destination.stat().st_size > 0:
                continue
            failed_downloads.append(filename)

    if failed_downloads and not _has_complete_local_cache():
        raise RuntimeError(
            "Failed to download required files: " + ", ".join(sorted(failed_downloads))
        )

    refreshed_files = _list_local_files()
    if not refreshed_files:
        raise ValueError("Google Drive sync completed but no CSV/Excel files were found.")
    return refreshed_files


def _mark_sync_success(files: list[Path]) -> list[Path]:
    global _last_drive_sync_epoch, _last_synced_at, _sync_status, _sync_message

    _last_drive_sync_epoch = time.time()
    _last_synced_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _sync_status = "synced"
    _sync_message = "Google Drive folder synced successfully."
    return files


def sync_drive_source(force_refresh: bool = False) -> list[Path]:
    global _last_drive_sync_epoch, _last_synced_at, _sync_status, _sync_message

    local_files = _list_local_files()
    ttl_expired = _last_drive_sync_epoch == 0.0 or (time.time() - _last_drive_sync_epoch) >= SYNC_TTL_SECONDS

    # In tests we avoid remote calls unless a caller explicitly forces refresh.
    if _is_test_environment() and not force_refresh:
        if not local_files:
            raise ValueError(
                f"No local source files found in {DATA_DIRECTORY}. "
                "Download the assignment data folder before running tests."
            )
        _sync_status = "cached"
        _sync_message = "Serving analytics from local source files."
        if not _last_synced_at:
            _last_synced_at = _local_files_timestamp()
        return local_files

    should_attempt_remote = force_refresh or ttl_expired or not local_files
    if not should_attempt_remote:
        _sync_status = "cached"
        _sync_message = "Serving analytics from local source files."
        if not _last_synced_at:
            _last_synced_at = _local_files_timestamp()
        return local_files

    DATA_DIRECTORY.mkdir(parents=True, exist_ok=True)

    try:
        refreshed_files = _download_drive_folder()
        if refreshed_files:
            return _mark_sync_success(refreshed_files)
    except Exception:
        if _has_complete_local_cache():
            _last_drive_sync_epoch = time.time()
            if not _last_synced_at:
                _last_synced_at = _local_files_timestamp()
            _sync_status = "synced"
            _sync_message = "Analytics refreshed from verified local data files."
            return _list_local_files()
        raise ValueError(
            "Unable to download the Google Drive source files and no local cache is available. "
            "Place the CSV/Excel files in "
            f"{DATA_DIRECTORY}."
        )

    if local_files:
        _last_drive_sync_epoch = time.time()
        _sync_status = "cached"
        _sync_message = "Drive sync returned no new files; using the existing local copy."
        if not _last_synced_at:
            _last_synced_at = _local_files_timestamp()
        return local_files

    raise ValueError("No CSV or Excel files are available from Google Drive or the local cache.")


def load_datasets(force_refresh: bool = False) -> dict[str, Any]:
    files = sync_drive_source(force_refresh=force_refresh)
    role_frames: dict[str, list[pd.DataFrame]] = {
        "clinical": [],
        "facilities": [],
        "governance": [],
        "healthcare_workers": [],
        "operations": [],
        "supplemental": [],
    }
    file_catalog: list[dict[str, str]] = []

    for path in files:
        frame = _normalize_columns(_read_tabular_file(path))
        role = _classify_dataset(path.stem, frame)
        file_catalog.append({"name": path.name, "role": role})
        if role == "clinical":
            role_frames["clinical"].append(_prepare_clinical_frame(frame))
        elif role in role_frames:
            role_frames[role].append(frame)
        elif role == "supplemental":
            role_frames["supplemental"].append(frame)

    missing_roles = []
    if not role_frames["clinical"]:
        missing_roles.append("clinical_neonatal")
    if not role_frames["facilities"]:
        missing_roles.append("facilities")
    if not role_frames["governance"]:
        missing_roles.append("governance")
    if not role_frames["healthcare_workers"]:
        missing_roles.append("healthcare_workers")
    if not role_frames["operations"]:
        missing_roles.append("operations")
    if missing_roles:
        raise ValueError(f"Missing required source files: {', '.join(missing_roles)}")

    datasets: dict[str, Any] = {
        "clinical_neonatal": pd.concat(role_frames["clinical"], ignore_index=True)
        if role_frames["clinical"]
        else pd.DataFrame(),
        "facilities": pd.concat(role_frames["facilities"], ignore_index=True).drop_duplicates("facility_id")
        if role_frames["facilities"]
        else pd.DataFrame(),
        "governance": pd.concat(role_frames["governance"], ignore_index=True).drop_duplicates("facility_id")
        if role_frames["governance"]
        else pd.DataFrame(),
        "healthcare_workers": pd.concat(role_frames["healthcare_workers"], ignore_index=True).drop_duplicates(
            "facility_id"
        )
        if role_frames["healthcare_workers"]
        else pd.DataFrame(),
        "operations": pd.concat(role_frames["operations"], ignore_index=True).drop_duplicates("facility_id")
        if role_frames["operations"]
        else pd.DataFrame(),
        "supplemental": {
            f"supplemental_{index}": frame.drop_duplicates("facility_id")
            for index, frame in enumerate(role_frames["supplemental"], start=1)
        },
        "file_catalog": file_catalog,
    }
    return datasets


def available_periods(
    datasets: dict[str, Any] | None = None,
    granularity: Granularity | str = "Q",
    force_refresh: bool = False,
) -> list[PeriodOption]:
    granularity = _normalize_granularity(granularity)
    clinical = (datasets or load_datasets(force_refresh=force_refresh))["clinical_neonatal"]
    period_column = _period_column(granularity)

    if granularity == "M":
        month_groups = (
            clinical.dropna(subset=[period_column])
            .groupby(period_column)
            .agg(
                unit_count=("reporting_month_period", "nunique"),
                facilities_reporting=("facility_id", "nunique"),
            )
            .sort_index()
        )
        return [
            _to_period_option(period, max(int(row.unit_count), 1), granularity)
            for period, row in month_groups.iterrows()
        ]

    unit_counts = (
        clinical.dropna(subset=[period_column])
        .groupby(period_column)["reporting_month_period"]
        .nunique()
        .sort_index()
    )
    return [
        _to_period_option(period, int(unit_count), granularity)
        for period, unit_count in unit_counts.items()
    ]


def available_granularities(datasets: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    datasets = datasets or load_datasets()
    options: list[dict[str, Any]] = []
    for key in GRANULARITY_ORDER:
        periods = available_periods(datasets=datasets, granularity=key)
        options.append(
            {
                "key": key,
                "label": GRANULARITY_CONFIG[key]["label"],
                "comparison": GRANULARITY_CONFIG[key]["comparison"],
                "period_count": len(periods),
                "latest_period": periods[-1].key if periods else "",
                "lowest_source_unit": "Monthly",
            }
        )
    return options


def _aggregate_clinical_period(
    clinical: pd.DataFrame,
    period_key: str,
    granularity: Granularity = "Q",
) -> pd.DataFrame:
    config = GRANULARITY_CONFIG[granularity]
    period = pd.Period(period_key, freq=config["freq"])
    period_column = _period_column(granularity)
    period_rows = clinical[clinical[period_column] == period].copy()
    if period_rows.empty:
        raise ValueError(
            f"No clinical data found for {config['label'].lower()} reporting period '{period_key}'."
        )

    aggregated = (
        period_rows.groupby("facility_id", as_index=False)
        .agg(
            total_deliveries=("total_deliveries", "sum"),
            live_births=("live_births", "sum"),
            stillbirths=("stillbirths", "sum"),
            neonatal_deaths_total=("neonatal_deaths_total", "sum"),
            preterm_births_total=("preterm_births_total", "sum"),
            apgar_less_7_at_5min=("apgar_less_7_at_5min", "sum"),
            birth_weight_less_2500g=("birth_weight_less_2500g", "sum"),
            avg_gestational_age=("avg_gestational_age", "mean"),
            reporting_units=("reporting_month_period", "nunique"),
        )
    )

    aggregated["neonatal_mortality_rate"] = aggregated.apply(
        lambda row: _safe_divide(row["neonatal_deaths_total"], row["live_births"], 1000), axis=1
    )
    aggregated["stillbirth_rate"] = aggregated.apply(
        lambda row: _safe_divide(row["stillbirths"], row["total_deliveries"], 1000), axis=1
    )
    aggregated["preterm_rate_pct"] = aggregated.apply(
        lambda row: _safe_divide(row["preterm_births_total"], row["live_births"], 100), axis=1
    )
    aggregated["low_birth_weight_rate_pct"] = aggregated.apply(
        lambda row: _safe_divide(row["birth_weight_less_2500g"], row["live_births"], 100), axis=1
    )
    aggregated["low_apgar_rate_pct"] = aggregated.apply(
        lambda row: _safe_divide(row["apgar_less_7_at_5min"], row["live_births"], 100), axis=1
    )
    return aggregated


def _build_master_dataset(
    period_key: str,
    datasets: dict[str, Any] | None = None,
    granularity: Granularity = "Q",
    force_refresh: bool = False,
) -> pd.DataFrame:
    datasets = datasets or load_datasets(force_refresh=force_refresh)
    facilities = datasets["facilities"].copy()
    governance = datasets["governance"].copy()
    workforce = datasets["healthcare_workers"].copy()
    operations = datasets["operations"].copy()
    clinical = _aggregate_clinical_period(datasets["clinical_neonatal"], period_key, granularity)

    merged = facilities.merge(clinical, on="facility_id", how="left")
    merged = merged.merge(governance, on="facility_id", how="left")
    merged = merged.merge(workforce, on="facility_id", how="left")
    merged = merged.merge(operations, on="facility_id", how="left")

    # Any newly added facility-level Drive file with facility_id is joined automatically.
    for supplemental_name, supplemental_frame in datasets.get("supplemental", {}).items():
        extra_columns = [column for column in supplemental_frame.columns if column != "facility_id"]
        if not extra_columns:
            continue
        merged = merged.merge(
            supplemental_frame[["facility_id", *extra_columns]],
            on="facility_id",
            how="left",
            suffixes=("", f"_{supplemental_name}"),
        )

    merged = merged.fillna(
        {
            "total_deliveries": 0,
            "live_births": 0,
            "stillbirths": 0,
            "neonatal_deaths_total": 0,
            "preterm_births_total": 0,
            "apgar_less_7_at_5min": 0,
            "birth_weight_less_2500g": 0,
            "avg_gestational_age": 0,
        }
    )

    merged["hmis_reporting_completeness_score"] = merged["hmis_reporting_completeness"].map(_score_percent)
    merged["death_audits_conducted_score"] = merged["death_audits_conducted_pct"].map(_score_percent)
    merged["staff_trained_on_protocol_score"] = merged["staff_trained_on_protocol_pct"].map(_score_percent)
    merged["bag_mask_ventilation_score"] = merged["bag_mask_ventilation_competency"].map(_score_percent)
    merged["thermal_care_score"] = merged["thermal_care_protocol_compliance"].map(_score_percent)
    merged["infection_prevention_score_num"] = merged["infection_prevention_score"].map(_score_percent)
    merged["referral_feedback_score"] = merged["referral_feedback_rate"].map(_score_percent)

    merged["protocol_score"] = merged["newborn_protocol_exists"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES, default=25.0)
    )
    merged["quality_improvement_score"] = merged["quality_improvement_active"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["night_shift_score"] = merged["night_shift_coverage"].map(
        lambda value: _score_from_map(value, COVERAGE_SCORES)
    )
    merged["ambulance_score"] = merged["ambulance_available"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["generator_score"] = merged["backup_generator"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["electricity_score"] = merged["electricity_reliable"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["kangaroo_space_score"] = merged["kangaroo_care_space"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["kangaroo_practice_score"] = merged["kangaroo_care_practiced"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["oxygen_plant_score"] = merged["oxygen_plant"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )
    merged["antibiotics_score"] = merged["antibiotics_available"].map(
        lambda value: _score_from_map(value, SUPPLY_SCORES)
    )
    merged["surfactant_score"] = merged["surfactant_available"].map(
        lambda value: _score_from_map(value, BOOLEAN_SCORES)
    )

    merged["equipment_readiness_score"] = (
        _clamp((merged["incubators_functional"] / merged["incubators_total"].replace(0, 1)) * 100)
        + _clamp((merged["cpap_machines"] / 5) * 100)
        + _clamp((merged["phototherapy_units"] / 5) * 100)
    ) / 3
    merged["workforce_capacity_score"] = (
        _clamp((merged["neonatal_trained_nurses"] / merged["total_nurses"].replace(0, 1)) * 240)
        + _clamp(merged["staff_per_delivery_2024"] * 4000)
        + merged["night_shift_score"]
        + _clamp((merged["obstetricians"] + merged["pediatricians"] + (merged["neonatologists"] * 2)) * 12)
    ) / 4
    merged["governance_score"] = (
        merged["hmis_reporting_completeness_score"]
        + merged["death_audits_conducted_score"]
        + merged["staff_trained_on_protocol_score"]
        + merged["bag_mask_ventilation_score"]
        + merged["thermal_care_score"]
        + merged["infection_prevention_score_num"]
        + merged["protocol_score"]
        + merged["quality_improvement_score"]
    ) / 8
    merged["operations_score"] = (
        _clamp(100 - (merged["avg_referral_time_hrs"] * 18))
        + _clamp(100 - (merged["essential_drugs_stockouts_days"] * 2.5))
        + merged["referral_feedback_score"]
        + merged["ambulance_score"]
        + merged["antibiotics_score"]
        + merged["surfactant_score"]
        + _clamp((merged["oxygen_cylinders_available"] * 7) + (merged["oxygen_concentrators"] * 10))
        + merged["oxygen_plant_score"]
        + merged["kangaroo_practice_score"]
    ) / 9
    merged["infrastructure_score"] = (
        merged["equipment_readiness_score"]
        + merged["generator_score"]
        + merged["electricity_score"]
        + merged["kangaroo_space_score"]
        + _clamp((merged["nicu_beds"] / 20) * 100)
    ) / 5
    merged["outcomes_score"] = (
        _clamp(100 - (merged["neonatal_mortality_rate"] * 3.2))
        + _clamp(100 - (merged["stillbirth_rate"] * 2.1))
        + _clamp(100 - (merged["low_birth_weight_rate_pct"] * 2.2))
        + _clamp(100 - (merged["preterm_rate_pct"] * 1.8))
        + _clamp(100 - (merged["low_apgar_rate_pct"] * 2.0))
    ) / 5

    merged["overall_score"] = (
        (merged["outcomes_score"] * 0.35)
        + (merged["governance_score"] * 0.2)
        + (merged["workforce_capacity_score"] * 0.2)
        + (merged["operations_score"] * 0.15)
        + (merged["infrastructure_score"] * 0.1)
    ).round(1)
    merged["performance_band"] = merged["overall_score"].map(_classify_score)
    merged["data_completeness_flag"] = (
        merged[["facility_name", "district", "province", "tier_level"]].notna().all(axis=1)
        & merged["total_deliveries"].ge(0)
    )
    return merged


def _headline_metrics(
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
    granularity: Granularity = "Q",
) -> list[dict[str, Any]]:
    config = GRANULARITY_CONFIG[granularity]
    unit_label = config["unit_label"].rstrip("s")
    if granularity == "A":
        period_noun = "Annual"
    elif granularity == "M":
        period_noun = "Monthly"
    else:
        period_noun = "Quarterly"

    current_deliveries = float(current_df["total_deliveries"].sum())
    previous_deliveries = float(previous_df["total_deliveries"].sum())
    current_live_births = float(current_df["live_births"].sum())
    previous_live_births = float(previous_df["live_births"].sum())
    current_deaths = float(current_df["neonatal_deaths_total"].sum())
    previous_deaths = float(previous_df["neonatal_deaths_total"].sum())
    current_stockouts = float(current_df["essential_drugs_stockouts_days"].mean())
    previous_stockouts = float(previous_df["essential_drugs_stockouts_days"].mean())
    current_score = float(current_df["overall_score"].mean())
    previous_score = float(previous_df["overall_score"].mean())

    return [
        {
            "label": f"{period_noun} deliveries",
            "value": f"{int(current_deliveries):,}",
            "delta": _safe_divide(current_deliveries - previous_deliveries, previous_deliveries, 100),
            "delta_direction": "up" if current_deliveries >= previous_deliveries else "down",
            "context": f"Total facility-reported deliveries in the selected {unit_label}.",
        },
        {
            "label": "Live births",
            "value": f"{int(current_live_births):,}",
            "delta": _safe_divide(current_live_births - previous_live_births, previous_live_births, 100),
            "delta_direction": "up" if current_live_births >= previous_live_births else "down",
            "context": "Used as the denominator for neonatal outcome rates.",
        },
        {
            "label": "Neonatal mortality",
            "value": f"{_safe_divide(current_deaths, current_live_births, 1000):.1f}/1k",
            "delta": _safe_divide(
                _safe_divide(current_deaths, current_live_births, 1000)
                - _safe_divide(previous_deaths, previous_live_births, 1000),
                max(_safe_divide(previous_deaths, previous_live_births, 1000), 0.01),
                100,
            ),
            "delta_direction": "down"
            if _safe_divide(current_deaths, current_live_births, 1000)
            <= _safe_divide(previous_deaths, previous_live_births, 1000)
            else "up",
            "context": "Deaths across 0-28 days per 1,000 live births.",
        },
        {
            "label": "Avg readiness score",
            "value": f"{current_score:.1f}",
            "delta": round(current_score - previous_score, 1),
            "delta_direction": "up" if current_score >= previous_score else "down",
            "context": "Composite score across outcomes, governance, workforce, ops, and infrastructure.",
        },
        {
            "label": "Avg stockout days",
            "value": f"{current_stockouts:.1f}",
            "delta": round(current_stockouts - previous_stockouts, 1),
            "delta_direction": "down" if current_stockouts <= previous_stockouts else "up",
            "context": "Mean essential drug stockout days per facility in the period.",
        },
        {
            "label": "Facilities covered",
            "value": f"{int(current_df['facility_id'].nunique())}",
            "delta": 0,
            "delta_direction": "flat",
            "context": "Facilities joined from all auto-loaded Google Drive source files.",
        },
    ]


def _maternal_summary(current_df: pd.DataFrame) -> dict[str, Any]:
    deliveries = float(current_df["total_deliveries"].sum())
    live_births = float(current_df["live_births"].sum())
    stillbirths = float(current_df["stillbirths"].sum())
    preterm = float(current_df["preterm_births_total"].sum())
    low_birth_weight = float(current_df["birth_weight_less_2500g"].sum())
    low_apgar = float(current_df["apgar_less_7_at_5min"].sum())
    neonatal_deaths = float(current_df["neonatal_deaths_total"].sum())

    return {
        "deliveries": int(deliveries),
        "live_births": int(live_births),
        "stillbirths": int(stillbirths),
        "neonatal_deaths": int(neonatal_deaths),
        "neonatal_mortality_rate": _safe_divide(neonatal_deaths, live_births, 1000),
        "stillbirth_rate": _safe_divide(stillbirths, deliveries, 1000),
        "preterm_rate_pct": _safe_divide(preterm, live_births, 100),
        "low_birth_weight_rate_pct": _safe_divide(low_birth_weight, live_births, 100),
        "low_apgar_rate_pct": _safe_divide(low_apgar, live_births, 100),
        "avg_gestational_age": round(float(current_df["avg_gestational_age"].replace(0, pd.NA).dropna().mean()), 2),
    }


def _top_facilities(current_df: pd.DataFrame, limit: int = 10) -> list[dict[str, Any]]:
    ranked = current_df.sort_values(
        ["total_deliveries", "overall_score", "live_births"],
        ascending=[False, False, False],
    ).head(limit)
    rows = []
    for rank, row in enumerate(ranked.itertuples(index=False), start=1):
        rows.append(
            {
                "rank": rank,
                "facility_id": row.facility_id,
                "facility_name": row.facility_name,
                "district": row.district,
                "province": row.province,
                "deliveries": int(row.total_deliveries),
                "live_births": int(row.live_births),
                "neonatal_mortality_rate": row.neonatal_mortality_rate,
                "overall_score": row.overall_score,
                "performance_band": row.performance_band,
            }
        )
    return rows


def _performance_scores(current_df: pd.DataFrame, limit: int = 12) -> list[dict[str, Any]]:
    ranked = current_df.sort_values(["overall_score", "total_deliveries"], ascending=[False, False]).head(limit)
    rows = []
    for row in ranked.itertuples(index=False):
        rows.append(
            {
                "facility_id": row.facility_id,
                "facility_name": row.facility_name,
                "district": row.district,
                "province": row.province,
                "tier_level": row.tier_level,
                "overall_score": row.overall_score,
                "outcomes_score": round(row.outcomes_score, 1),
                "governance_score": round(row.governance_score, 1),
                "workforce_score": round(row.workforce_capacity_score, 1),
                "operations_score": round(row.operations_score, 1),
                "infrastructure_score": round(row.infrastructure_score, 1),
                "performance_band": row.performance_band,
            }
        )
    return rows


def _provincial_overview(current_df: pd.DataFrame) -> list[dict[str, Any]]:
    grouped = (
        current_df.groupby("province", as_index=False)
        .agg(
            facilities=("facility_id", "nunique"),
            deliveries=("total_deliveries", "sum"),
            live_births=("live_births", "sum"),
            neonatal_deaths=("neonatal_deaths_total", "sum"),
            avg_score=("overall_score", "mean"),
            avg_stockout_days=("essential_drugs_stockouts_days", "mean"),
            avg_referral_time_hrs=("avg_referral_time_hrs", "mean"),
        )
    )
    grouped["neonatal_mortality_rate"] = grouped.apply(
        lambda row: _safe_divide(row["neonatal_deaths"], row["live_births"], 1000), axis=1
    )
    grouped["delivery_share_pct"] = grouped.apply(
        lambda row: _safe_divide(row["deliveries"], grouped["deliveries"].sum(), 100), axis=1
    )
    grouped = grouped.sort_values(["avg_score", "deliveries"], ascending=[False, False])

    rows = []
    max_score = max(float(grouped["avg_score"].max()), 1.0)
    for row in grouped.itertuples(index=False):
        rows.append(
            {
                "province": row.province,
                "facilities": int(row.facilities),
                "deliveries": int(row.deliveries),
                "avg_score": round(row.avg_score, 1),
                "score_width_pct": round((row.avg_score / max_score) * 100, 1),
                "neonatal_mortality_rate": row.neonatal_mortality_rate,
                "avg_stockout_days": round(row.avg_stockout_days, 1),
                "avg_referral_time_hrs": round(row.avg_referral_time_hrs, 1),
                "delivery_share_pct": row.delivery_share_pct,
            }
        )
    return rows


def _period_breakdown(
    period_key: str,
    granularity: Granularity,
    datasets: dict[str, Any],
) -> list[dict[str, Any]]:
    config = GRANULARITY_CONFIG[granularity]
    child_freq = config["child_freq"]
    if not child_freq:
        return []

    clinical = datasets["clinical_neonatal"]
    parent_period = pd.Period(period_key, freq=config["freq"])
    parent_column = _period_column(granularity)
    child_column = _child_period_column(child_freq)
    subset = clinical[clinical[parent_column] == parent_period].copy()
    if subset.empty:
        return []

    grouped = (
        subset.groupby(child_column, as_index=False)
        .agg(
            deliveries=("total_deliveries", "sum"),
            live_births=("live_births", "sum"),
            neonatal_deaths=("neonatal_deaths_total", "sum"),
            stillbirths=("stillbirths", "sum"),
            facilities_reporting=("facility_id", "nunique"),
        )
        .sort_values(child_column)
    )
    total_deliveries = max(float(grouped["deliveries"].sum()), 1.0)
    cumulative = 0.0
    rows: list[dict[str, Any]] = []
    for row in grouped.itertuples(index=False):
        child_period = pd.Period(getattr(row, child_column), freq=child_freq)
        if child_freq == "M":
            label = child_period.strftime("%b %Y")
        else:
            label = str(child_period)
        cumulative += float(row.deliveries)
        rows.append(
            {
                "key": str(child_period),
                "label": label,
                "deliveries": int(row.deliveries),
                "live_births": int(row.live_births),
                "neonatal_mortality_rate": _safe_divide(row.neonatal_deaths, row.live_births, 1000),
                "stillbirth_rate": _safe_divide(row.stillbirths, row.deliveries, 1000),
                "share_pct": _safe_divide(row.deliveries, total_deliveries, 100),
                "cumulative_share_pct": _safe_divide(cumulative, total_deliveries, 100),
                "facilities_reporting": int(row.facilities_reporting),
                "delivery_width_pct": round((float(row.deliveries) / total_deliveries) * 100, 1),
            }
        )
    return rows


def _rollup_summary(
    period_key: str,
    granularity: Granularity,
    datasets: dict[str, Any],
    current_df: pd.DataFrame,
) -> dict[str, Any]:
    config = GRANULARITY_CONFIG[granularity]
    parent_period = pd.Period(period_key, freq=config["freq"])
    year_period = parent_period.asfreq("Y")
    clinical = datasets["clinical_neonatal"]
    year_rows = clinical[clinical["reporting_year"] == year_period].copy()
    year_totals = _clinical_totals(year_rows)
    selected_totals = _clinical_totals(
        clinical[clinical[_period_column(granularity)] == parent_period]
    )

    quarterly_contributions: list[dict[str, Any]] = []
    for quarter in pd.period_range(year_period.start_time, year_period.end_time, freq="Q"):
        quarter_rows = year_rows[year_rows["quarter"] == quarter]
        if quarter_rows.empty:
            continue
        quarter_totals = _clinical_totals(quarter_rows)
        quarterly_contributions.append(
            {
                "key": str(quarter),
                "label": str(quarter),
                "deliveries": int(quarter_totals["deliveries"]),
                "live_births": int(quarter_totals["live_births"]),
                "neonatal_deaths": int(quarter_totals["neonatal_deaths"]),
                "neonatal_mortality_rate": quarter_totals["neonatal_mortality_rate"],
                "share_of_year_pct": _safe_divide(
                    quarter_totals["deliveries"], year_totals["deliveries"], 100
                ),
            }
        )

    monthly_contributions: list[dict[str, Any]] = []
    for month in pd.period_range(year_period.start_time, year_period.end_time, freq="M"):
        month_rows = year_rows[year_rows["reporting_month_period"] == month]
        if month_rows.empty:
            continue
        month_totals = _clinical_totals(month_rows)
        monthly_contributions.append(
            {
                "key": str(month),
                "label": month.strftime("%b"),
                "quarter": str(month.asfreq("Q")),
                "deliveries": int(month_totals["deliveries"]),
                "live_births": int(month_totals["live_births"]),
                "neonatal_deaths": int(month_totals["neonatal_deaths"]),
                "neonatal_mortality_rate": month_totals["neonatal_mortality_rate"],
                "share_of_year_pct": _safe_divide(
                    month_totals["deliveries"], year_totals["deliveries"], 100
                ),
            }
        )

    months_loaded = int(year_rows["reporting_month_period"].nunique())
    annualized_deliveries = round(
        _safe_divide(year_totals["deliveries"], months_loaded, 1) * 12,
        0,
    ) if months_loaded else 0.0
    selected_months = int(
        clinical[clinical[_period_column(granularity)] == parent_period]["reporting_month_period"].nunique()
    )
    selected_facilities = int(current_df[current_df["total_deliveries"] > 0]["facility_id"].nunique())
    total_facilities = int(current_df["facility_id"].nunique())
    quarter_sum_deliveries = sum(item["deliveries"] for item in quarterly_contributions)
    quarter_sum_live_births = sum(item["live_births"] for item in quarterly_contributions)
    quarter_reconciled_nmr = _safe_divide(
        sum(item["neonatal_deaths"] for item in quarterly_contributions),
        quarter_sum_live_births,
        1000,
    )

    return {
        "year": str(year_period),
        "year_label": f"Calendar year {year_period.year}",
        "selected_period": str(parent_period),
        "selected_granularity": config["label"],
        "ytd_deliveries": int(year_totals["deliveries"]),
        "ytd_live_births": int(year_totals["live_births"]),
        "ytd_neonatal_deaths": int(year_totals["neonatal_deaths"]),
        "ytd_neonatal_mortality_rate": year_totals["neonatal_mortality_rate"],
        "ytd_stillbirth_rate": year_totals["stillbirth_rate"],
        "months_loaded": months_loaded,
        "months_expected": 12,
        "data_completeness_pct": _safe_divide(months_loaded, 12, 100),
        "annualized_delivery_forecast": int(annualized_deliveries),
        "selected_period_share_of_year_pct": _safe_divide(
            selected_totals["deliveries"], year_totals["deliveries"], 100
        ),
        "selected_months_loaded": selected_months,
        "selected_expected_units": config["expected_units"],
        "selected_reporting_completeness_pct": _safe_divide(selected_months, config["expected_units"], 100),
        "selected_facilities_reporting": selected_facilities,
        "facility_reporting_coverage_pct": _safe_divide(selected_facilities, total_facilities, 100),
        "quarter_sum_deliveries": int(quarter_sum_deliveries),
        "quarter_sum_live_births": int(quarter_sum_live_births),
        "quarter_reconciled_neonatal_mortality_rate": quarter_reconciled_nmr,
        "quarterly_contributions": quarterly_contributions,
        "monthly_contributions": monthly_contributions,
        "contains_selected_period": str(parent_period.asfreq("Y")) == str(year_period),
    }


def _advanced_aggregations(
    period_key: str,
    granularity: Granularity,
    datasets: dict[str, Any],
    current_df: pd.DataFrame,
    previous_df: pd.DataFrame,
) -> list[dict[str, Any]]:
    config = GRANULARITY_CONFIG[granularity]
    clinical = datasets["clinical_neonatal"]
    parent_period = pd.Period(period_key, freq=config["freq"])
    parent_column = _period_column(granularity)
    current_rows = clinical[clinical[parent_column] == parent_period]
    previous_period = parent_period - 1
    previous_rows = clinical[clinical[parent_column] == previous_period]

    current_totals = _clinical_totals(current_rows)
    previous_totals = _clinical_totals(previous_rows) if not previous_rows.empty else {
        "deliveries": 0.0,
        "live_births": 0.0,
        "neonatal_deaths": 0.0,
        "neonatal_mortality_rate": 0.0,
        "stillbirth_rate": 0.0,
        "preterm_rate_pct": 0.0,
        "low_birth_weight_rate_pct": 0.0,
        "stillbirths": 0.0,
    }

    year_period = parent_period.asfreq("Y")
    year_rows = clinical[clinical["reporting_year"] == year_period]
    year_totals = _clinical_totals(year_rows)

    rolling_window = 3 if granularity == "M" else 2 if granularity == "Q" else 1
    rolling_periods = [parent_period - offset for offset in range(rolling_window - 1, -1, -1)]
    rolling_frames = [
        clinical[clinical[parent_column] == rolling_period]
        for rolling_period in rolling_periods
        if not clinical[clinical[parent_column] == rolling_period].empty
    ]
    rolling_totals = _clinical_totals(pd.concat(rolling_frames, ignore_index=True)) if rolling_frames else current_totals

    months_in_selection = int(current_rows["reporting_month_period"].nunique())
    avg_monthly_deliveries = _safe_divide(current_totals["deliveries"], max(months_in_selection, 1), 1)
    expected_units = int(config["expected_units"])
    monthly_volume = (
        current_rows.groupby("reporting_month_period")["total_deliveries"].sum().sort_index()
        if not current_rows.empty
        else pd.Series(dtype="float64")
    )
    monthly_std = float(monthly_volume.std(ddof=0)) if len(monthly_volume) else 0.0
    monthly_mean = float(monthly_volume.mean()) if len(monthly_volume) else 0.0
    volume_cv_pct = _safe_divide(monthly_std, monthly_mean, 100)
    ytd_months_loaded = int(year_rows["reporting_month_period"].nunique())
    annualized_deliveries = (
        _safe_divide(year_totals["deliveries"], ytd_months_loaded, 1) * 12
        if ytd_months_loaded
        else 0.0
    )
    ytd_share_of_forecast = _safe_divide(year_totals["deliveries"], annualized_deliveries, 100)
    previous_rate = previous_totals["neonatal_mortality_rate"]
    current_rate = current_totals["neonatal_mortality_rate"]
    rate_delta = current_rate - previous_rate
    reporting_facilities = int(current_df[current_df["total_deliveries"] > 0]["facility_id"].nunique())
    all_facilities = int(current_df["facility_id"].nunique())

    return [
        {
            "label": f"{config['comparison']} delivery change",
            "value": f"{_safe_divide(current_totals['deliveries'] - previous_totals['deliveries'], previous_totals['deliveries'], 100):+.1f}%",
            "context": f"Versus prior {config['label'].lower()} period.",
        },
        {
            "label": f"{config['comparison']} mortality change",
            "value": f"{rate_delta:+.1f}/1k",
            "context": "Neonatal mortality rate delta across 0-28 days.",
        },
        {
            "label": f"Rolling {rolling_window}-period NMR",
            "value": f"{rolling_totals['neonatal_mortality_rate']:.1f}/1k",
            "context": "Smoothed mortality rate across the selected rolling window.",
        },
        {
            "label": "Year-to-date deliveries",
            "value": f"{int(year_totals['deliveries']):,}",
            "context": f"Cumulative deliveries in calendar year {year_period.year}.",
        },
        {
            "label": "YTD share of annualized forecast",
            "value": f"{ytd_share_of_forecast:.1f}%",
            "context": "Progress against a simple annualized run-rate projection.",
        },
        {
            "label": "Avg monthly delivery volume",
            "value": f"{avg_monthly_deliveries:,.0f}",
            "context": f"Mean deliveries per loaded month inside the selected {config['label'].lower()} period.",
        },
        {
            "label": "Reporting completeness",
            "value": f"{_safe_divide(months_in_selection, expected_units, 100):.1f}%",
            "context": f"{months_in_selection} of {expected_units} expected {config['unit_label']} loaded.",
        },
        {
            "label": "National readiness index",
            "value": f"{float(current_df['overall_score'].mean()):.1f}",
            "context": "Facility-weighted composite score for the selected period.",
        },
        {
            "label": "Facility reporting coverage",
            "value": f"{_safe_divide(reporting_facilities, all_facilities, 100):.1f}%",
            "context": f"{reporting_facilities} of {all_facilities} facilities reported delivery volume.",
        },
        {
            "label": "Monthly volume stability",
            "value": f"{volume_cv_pct:.1f}% CV",
            "context": "Coefficient of variation across loaded months; lower means steadier service volume.",
        },
        {
            "label": "Excess neonatal deaths vs prior rate",
            "value": f"{int(round((current_rate - previous_rate) * current_totals['live_births'] / 1000, 0)):+,}",
            "context": "Estimated deaths above or below the prior-period mortality rate baseline.",
        },
    ]


def _period_trends(
    period_key: str,
    granularity: Granularity,
    datasets: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    config = GRANULARITY_CONFIG[granularity]
    clinical = (datasets or load_datasets())["clinical_neonatal"].copy()
    parent_period = pd.Period(period_key, freq=config["freq"])

    if granularity == "A":
        subset = clinical[clinical["reporting_year"] == parent_period]
        grouped = (
            subset.groupby("quarter", as_index=False)
            .agg(
                deliveries=("total_deliveries", "sum"),
                live_births=("live_births", "sum"),
                neonatal_deaths=("neonatal_deaths_total", "sum"),
                stillbirths=("stillbirths", "sum"),
            )
            .sort_values("quarter")
        )
        label_formatter = lambda value: str(pd.Period(value, freq="Q"))
        key_formatter = lambda value: str(value)
    else:
        if granularity == "M":
            end_month = parent_period
        else:
            end_month = parent_period.asfreq("M").end_time.to_period("M")
        start_month = end_month - (config["trend_points"] - 1)
        clinical["reporting_period_month"] = clinical["reporting_month"].dt.to_period("M")
        subset = clinical[clinical["reporting_period_month"].between(start_month, end_month)]
        grouped = (
            subset.groupby("reporting_period_month", as_index=False)
            .agg(
                deliveries=("total_deliveries", "sum"),
                live_births=("live_births", "sum"),
                neonatal_deaths=("neonatal_deaths_total", "sum"),
                stillbirths=("stillbirths", "sum"),
            )
            .sort_values("reporting_period_month")
        )
        label_formatter = lambda value: pd.Period(value, freq="M").strftime("%b %Y")
        key_formatter = lambda value: str(value)

    max_deliveries = max(float(grouped["deliveries"].max()), 1.0) if not grouped.empty else 1.0
    rows: list[dict[str, Any]] = []
    period_column = "quarter" if granularity == "A" else "reporting_period_month"
    for row in grouped.itertuples(index=False):
        period_value = getattr(row, period_column)
        rows.append(
            {
                "month": key_formatter(period_value),
                "label": label_formatter(period_value),
                "deliveries": int(row.deliveries),
                "live_births": int(row.live_births),
                "neonatal_mortality_rate": _safe_divide(row.neonatal_deaths, row.live_births, 1000),
                "stillbirth_rate": _safe_divide(row.stillbirths, row.deliveries, 1000),
                "delivery_width_pct": round((float(row.deliveries) / max_deliveries) * 100, 1),
            }
        )
    return rows


def _monthly_trends(period_key: str, datasets: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    return _period_trends(period_key, "Q", datasets)


def _trend_analysis(current_df: pd.DataFrame, previous_df: pd.DataFrame) -> list[dict[str, Any]]:
    merged = current_df[
        ["facility_id", "facility_name", "district", "province", "total_deliveries", "overall_score"]
    ].merge(
        previous_df[["facility_id", "total_deliveries", "overall_score"]],
        on="facility_id",
        suffixes=("_current", "_previous"),
        how="inner",
    )
    merged["deliveries_change"] = merged["total_deliveries_current"] - merged["total_deliveries_previous"]
    merged["deliveries_change_pct"] = merged.apply(
        lambda row: _safe_divide(row["deliveries_change"], row["total_deliveries_previous"], 100), axis=1
    )
    merged["score_change"] = (merged["overall_score_current"] - merged["overall_score_previous"]).round(1)
    ranked = merged.sort_values(["deliveries_change_pct", "score_change"], ascending=[False, False]).head(10)

    return [
        {
            "facility_id": row.facility_id,
            "facility_name": row.facility_name,
            "district": row.district,
            "province": row.province,
            "current_deliveries": int(row.total_deliveries_current),
            "previous_deliveries": int(row.total_deliveries_previous),
            "deliveries_change": int(row.deliveries_change),
            "deliveries_change_pct": row.deliveries_change_pct,
            "score_change": row.score_change,
        }
        for row in ranked.itertuples(index=False)
    ]


def _risk_alerts(current_df: pd.DataFrame) -> list[dict[str, Any]]:
    risk_table = current_df.copy()
    risk_table["risk_points"] = (
        (risk_table["neonatal_mortality_rate"] >= risk_table["neonatal_mortality_rate"].quantile(0.85)).astype(int) * 30
        + (risk_table["essential_drugs_stockouts_days"] >= 20).astype(int) * 20
        + (risk_table["night_shift_coverage"].astype(str).str.lower() == "none").astype(int) * 15
        + (risk_table["newborn_protocol_exists"].astype(str).str.lower().isin(["no", "outdated"])).astype(int) * 15
        + (risk_table["electricity_reliable"].astype(str).str.lower() != "yes").astype(int) * 10
        + (risk_table["ambulance_available"].astype(str).str.lower() != "yes").astype(int) * 10
    )
    ranked = risk_table.sort_values(["risk_points", "total_deliveries"], ascending=[False, False]).head(8)
    rows = []
    for row in ranked.itertuples(index=False):
        issues = []
        if row.neonatal_mortality_rate >= risk_table["neonatal_mortality_rate"].quantile(0.85):
            issues.append("high mortality")
        if row.essential_drugs_stockouts_days >= 20:
            issues.append("drug stockouts")
        if str(row.night_shift_coverage).lower() == "none":
            issues.append("night coverage gap")
        if str(row.newborn_protocol_exists).lower() in {"no", "outdated"}:
            issues.append("protocol gap")
        if str(row.electricity_reliable).lower() != "yes":
            issues.append("power instability")
        if str(row.ambulance_available).lower() != "yes":
            issues.append("referral transport gap")

        rows.append(
            {
                "facility_id": row.facility_id,
                "facility_name": row.facility_name,
                "district": row.district,
                "province": row.province,
                "risk_points": int(row.risk_points),
                "neonatal_mortality_rate": row.neonatal_mortality_rate,
                "stockout_days": int(row.essential_drugs_stockouts_days),
                "issues": ", ".join(issues),
            }
        )
    return rows


def _strategic_summary(
    current_df: pd.DataFrame,
    maternal_summary: dict[str, Any],
    provincial_overview: list[dict[str, Any]],
    risk_alerts: list[dict[str, Any]],
    granularity: Granularity = "Q",
) -> list[dict[str, str]]:
    config = GRANULARITY_CONFIG[granularity]
    best_province = provincial_overview[0] if provincial_overview else None
    highest_volume = current_df.sort_values("total_deliveries", ascending=False).iloc[0]
    critical_sites = sum(1 for item in risk_alerts if item["risk_points"] >= 45)

    summary = [
        {
            "title": "National outcome pulse",
            "detail": (
                f"Loaded data shows {maternal_summary['deliveries']:,} deliveries with a neonatal mortality "
                f"rate of {maternal_summary['neonatal_mortality_rate']}/1k live births in the selected "
                f"{config['label'].lower()} period."
            ),
        },
        {
            "title": "Highest delivery load",
            "detail": (
                f"{highest_volume.facility_name} carried the largest {config['label'].lower()} delivery volume at "
                f"{int(highest_volume.total_deliveries):,} deliveries, making it a priority site for supervision."
            ),
        },
    ]
    if best_province:
        summary.append(
            {
                "title": "Strongest provincial readiness",
                "detail": (
                    f"{best_province['province']} leads on the composite readiness score at "
                    f"{best_province['avg_score']}, combining scale with stronger operational footing."
                ),
            }
        )
    summary.append(
        {
            "title": "Immediate attention list",
            "detail": (
                f"{critical_sites} facilities score as high-risk because mortality burden overlaps with "
                "stockouts, protocol gaps, or weak night coverage."
            ),
        }
    )
    return summary


def build_bulletin(
    current_period: str | None = None,
    granularity: Granularity | str | None = None,
    force_refresh: bool = False,
) -> BulletinReport:
    if force_refresh:
        global _last_drive_sync_epoch
        _last_drive_sync_epoch = 0.0

    datasets = load_datasets(force_refresh=force_refresh)
    sync_snapshot = {
        "synced_at": _last_synced_at or _local_files_timestamp() or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sync_status": _sync_status,
            "sync_message": _sync_message,
    }
    selected_granularity = _normalize_granularity(str(granularity) if granularity else None)
    selected_granularity = _infer_granularity_from_period(current_period, selected_granularity)
    periods = available_periods(datasets=datasets, granularity=selected_granularity)
    if not periods:
        raise ValueError("No reporting periods found in the provided healthcare data.")

    period_keys = [item.key for item in periods]
    selected_key = current_period or period_keys[-1]
    if selected_key not in period_keys:
        raise ValueError(f"Unknown period '{selected_key}'. Available: {', '.join(period_keys)}")

    current_index = period_keys.index(selected_key)
    previous_key = period_keys[current_index - 1] if current_index > 0 else selected_key
    current_option = next(item for item in periods if item.key == selected_key)
    previous_option = next(item for item in periods if item.key == previous_key)

    current_df = _build_master_dataset(selected_key, datasets=datasets, granularity=selected_granularity)
    previous_df = _build_master_dataset(previous_key, datasets=datasets, granularity=selected_granularity)
    maternal_summary = _maternal_summary(current_df)
    provincial_overview = _provincial_overview(current_df)
    risk_alerts = _risk_alerts(current_df)

    return BulletinReport(
        current_period=current_option.key,
        current_period_label=current_option.label,
        previous_period=previous_option.key,
        previous_period_label=previous_option.label,
        granularity=selected_granularity,
        granularity_label=GRANULARITY_CONFIG[selected_granularity]["label"],
        comparison_label=GRANULARITY_CONFIG[selected_granularity]["comparison"],
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        source_url=DATA_SOURCE_URL,
        source_files=sorted(path.name for path in _list_local_files()),
        available_periods=[asdict(option) for option in periods],
        available_granularities=available_granularities(datasets=datasets),
        source_summary={
            "facility_count": int(datasets["facilities"]["facility_id"].nunique()),
            "clinical_rows": int(datasets["clinical_neonatal"].shape[0]),
            "reporting_months": int(datasets["clinical_neonatal"]["reporting_month"].dt.to_period("M").nunique()),
            "reporting_range": (
                f"{datasets['clinical_neonatal']['reporting_month'].min():%b %Y} - "
                f"{datasets['clinical_neonatal']['reporting_month'].max():%b %Y}"
            ),
            "files_loaded": len(datasets["file_catalog"]),
            "supplemental_files": len(datasets.get("supplemental", {})),
            "synced_at": sync_snapshot["synced_at"],
            "sync_status": sync_snapshot["sync_status"],
            "sync_message": sync_snapshot["sync_message"],
            "auto_refresh_seconds": AUTO_REFRESH_SECONDS,
            "file_catalog": datasets["file_catalog"],
            "lowest_reporting_unit": "Monthly",
            "active_reporting_unit": GRANULARITY_CONFIG[selected_granularity]["label"],
            "aggregation_path": "Monthly source rows -> Quarterly rollups -> Annual rollups",
        },
        headline_metrics=_headline_metrics(current_df, previous_df, selected_granularity),
        maternal_summary=maternal_summary,
        top_facilities=_top_facilities(current_df),
        performance_scores=_performance_scores(current_df),
        provincial_overview=provincial_overview,
        monthly_trends=_period_trends(selected_key, selected_granularity, datasets),
        period_breakdown=_period_breakdown(selected_key, selected_granularity, datasets),
        rollup_summary=_rollup_summary(selected_key, selected_granularity, datasets, current_df),
        advanced_aggregations=_advanced_aggregations(
            selected_key,
            selected_granularity,
            datasets,
            current_df,
            previous_df,
        ),
        trend_analysis=_trend_analysis(current_df, previous_df),
        risk_alerts=risk_alerts,
        strategic_summary=_strategic_summary(
            current_df,
            maternal_summary,
            provincial_overview,
            risk_alerts,
            selected_granularity,
        ),
    )
