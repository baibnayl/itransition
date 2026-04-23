from __future__ import annotations

import base64
import io
import math
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from scipy.stats import t as student_t

try:
    from weasyprint import HTML
except Exception:
    HTML = None


app = Flask(__name__)

DEFAULT_SHEET_URL = os.getenv("GOOGLE_SHEET_CSV_URL", "https://docs.google.com/spreadsheets/d/1HqySGtl5cgkm7q8V249C4zSRizB5UV9BWTlZX0HS0gE/export?format=csv&gid=134185645").strip()
DEFAULT_IQR_K = float(os.getenv("DEFAULT_IQR_K", "1.5"))
DEFAULT_Z_THRESHOLD = float(os.getenv("DEFAULT_Z_THRESHOLD", "3.0"))
DEFAULT_MA_WINDOW = int(os.getenv("DEFAULT_MA_WINDOW", "7"))
DEFAULT_MA_PCT = float(os.getenv("DEFAULT_MA_PCT", "25.0"))
DEFAULT_GRUBBS_ALPHA = float(os.getenv("DEFAULT_GRUBBS_ALPHA", "0.05"))
DEFAULT_TREND_DEGREE = int(os.getenv("DEFAULT_TREND_DEGREE", "2"))


DATE_ALIASES = ["date", "day", "timestamp", "record_date", "report_date"]
MINE_ALIASES = ["mine", "mine_name", "site", "shaft", "unit", "location"]
OUTPUT_ALIASES = [
    "output",
    "daily_output",
    "final_output",
    "production",
    "value",
    "tons",
    "tonnage",
]


@dataclass
class DashboardParams:
    sheet_url: str
    chart_type: str = "line"
    trend_degree: int = DEFAULT_TREND_DEGREE
    iqr_k: float = DEFAULT_IQR_K
    z_threshold: float = DEFAULT_Z_THRESHOLD
    ma_window: int = DEFAULT_MA_WINDOW
    ma_pct: float = DEFAULT_MA_PCT
    grubbs_alpha: float = DEFAULT_GRUBBS_ALPHA
    entities: Optional[List[str]] = None


def safe_float(value: Any, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def safe_int(value: Any, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normalize_chart_type(value: str) -> str:
    value = (value or "line").strip().lower()
    return value if value in {"line", "bar", "stacked"} else "line"


def parse_entities(raw: str) -> Optional[List[str]]:
    if not raw:
        return None
    items = [x.strip() for x in raw.split(",") if x.strip()]
    return items or None


def parse_params() -> DashboardParams:
    payload = request.get_json(silent=True) or {}
    args = request.args

    sheet_url = (
        (payload.get("sheet_url") or args.get("sheet_url") or DEFAULT_SHEET_URL).strip()
    )
    params = DashboardParams(
        sheet_url=sheet_url,
        chart_type=normalize_chart_type(payload.get("chart_type") or args.get("chart_type")),
        trend_degree=max(1, min(4, safe_int(payload.get("trend_degree") or args.get("trend_degree"), DEFAULT_TREND_DEGREE))),
        iqr_k=max(0.1, safe_float(payload.get("iqr_k") or args.get("iqr_k"), DEFAULT_IQR_K)),
        z_threshold=max(0.1, safe_float(payload.get("z_threshold") or args.get("z_threshold"), DEFAULT_Z_THRESHOLD)),
        ma_window=max(2, safe_int(payload.get("ma_window") or args.get("ma_window"), DEFAULT_MA_WINDOW)),
        ma_pct=max(0.1, safe_float(payload.get("ma_pct") or args.get("ma_pct"), DEFAULT_MA_PCT)),
        grubbs_alpha=min(0.5, max(0.0001, safe_float(payload.get("grubbs_alpha") or args.get("grubbs_alpha"), DEFAULT_GRUBBS_ALPHA))),
        entities=parse_entities(payload.get("entities") or args.get("entities", "")),
    )
    return params


def normalize_google_sheet_url(url: str) -> str:
    url = (url or "").strip()
    if not url:
        raise ValueError("Google Sheet URL is empty. Pass ?sheet_url=... or set GOOGLE_SHEET_CSV_URL.")

    if "docs.google.com/spreadsheets" not in url:
        return url

    if "output=csv" in url or "format=csv" in url:
        return url

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("Could not parse Google Sheet document ID from the URL.")

    sheet_id = match.group(1)
    gid_match = re.search(r"[#&?]gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def _find_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    exact = {c.lower(): c for c in columns}
    for alias in aliases:
        if alias in exact:
            return exact[alias]
    for col in columns:
        c = col.lower()
        for alias in aliases:
            if alias in c:
                return col
    return None


def _to_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace(r"[^\d,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
        .replace({"": np.nan, "-": np.nan, ".": np.nan})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def load_dataset(sheet_url: str) -> pd.DataFrame:
    csv_url = normalize_google_sheet_url(sheet_url)
    raw = pd.read_csv(csv_url)
    raw.columns = [str(c).strip() for c in raw.columns]
    if raw.empty:
        raise ValueError("The Google Sheet is empty.")

    date_col = _find_column(raw.columns.tolist(), DATE_ALIASES)
    if not date_col:
        raise ValueError(
            "Could not detect a date column. Expected something like: date, day, timestamp."
        )

    mine_col = _find_column(raw.columns.tolist(), MINE_ALIASES)
    output_col = _find_column(raw.columns.tolist(), OUTPUT_ALIASES)

    if mine_col and output_col:
        df = raw[[date_col, mine_col, output_col]].copy()
        df.columns = ["date", "mine", "output"]
        df["output"] = _to_numeric(df["output"])
    else:
        numeric_candidates = []
        for col in raw.columns:
            if col == date_col:
                continue
            numeric_values = _to_numeric(raw[col])
            ratio = float(numeric_values.notna().mean()) if len(numeric_values) else 0.0
            if ratio >= 0.6:
                numeric_candidates.append(col)

        if len(numeric_candidates) < 1:
            raise ValueError(
                "Could not detect mine/output columns. Use long format [date, mine, output] "
                "or wide format [date, mineA, mineB, ...]."
            )

        df = raw[[date_col] + numeric_candidates].copy()
        df = df.melt(id_vars=[date_col], var_name="mine", value_name="output")
        df.columns = ["date", "mine", "output"]
        df["output"] = _to_numeric(df["output"])

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["mine"] = df["mine"].astype(str).str.strip()
    df = df.dropna(subset=["date", "mine", "output"]).copy()

    if df.empty:
        raise ValueError("No valid rows remained after parsing date/mine/output columns.")

    df["date"] = df["date"].dt.floor("D")
    df = (
        df.groupby(["date", "mine"], as_index=False)["output"]
        .sum()
        .sort_values(["date", "mine"])
        .reset_index(drop=True)
    )
    return df


def add_total_series(df: pd.DataFrame) -> pd.DataFrame:
    total = df.groupby("date", as_index=False)["output"].sum()
    total["mine"] = "Total"
    combined = pd.concat([df, total], ignore_index=True, axis=0)
    combined = combined.sort_values(["mine", "date"]).reset_index(drop=True)
    return combined


def compute_summary_stats(df: pd.DataFrame) -> pd.DataFrame:
    def _summ(group: pd.DataFrame) -> pd.Series:
        s = group["output"]
        q1 = float(s.quantile(0.25))
        q3 = float(s.quantile(0.75))
        iqr = q3 - q1
        return pd.Series(
            {
                "mean_daily_output": float(s.mean()),
                "std_dev": float(s.std(ddof=1)) if len(s) > 1 else 0.0,
                "median": float(s.median()),
                "q1": q1,
                "q3": q3,
                "iqr": iqr,
                "min_output": float(s.min()),
                "max_output": float(s.max()),
                "days_count": int(s.shape[0]),
            }
        )

    rows = []
    for mine_name, group in df.groupby("mine", sort=False):
        row = _summ(group).to_dict()
        row["mine"] = mine_name
        rows.append(row)
    return pd.DataFrame(rows)


def grubbs_iterative_flags(values: pd.Series, alpha: float) -> List[Dict[str, Any]]:
    arr = values.astype(float).to_numpy()
    original_indices = list(values.index)
    flagged: List[Dict[str, Any]] = []
    pass_no = 1

    while len(arr) >= 3:
        n = len(arr)
        mean = float(np.mean(arr))
        std = float(np.std(arr, ddof=1))
        if std == 0 or math.isnan(std):
            break

        deviations = np.abs(arr - mean)
        max_pos = int(np.argmax(deviations))
        candidate = float(arr[max_pos])
        g_stat = float(deviations[max_pos] / std)

        t_crit = float(student_t.ppf(1 - alpha / (2 * n), n - 2))
        g_crit = ((n - 1) / math.sqrt(n)) * math.sqrt((t_crit ** 2) / (n - 2 + t_crit ** 2))

        if g_stat > g_crit:
            flagged.append(
                {
                    "index": int(original_indices[max_pos]),
                    "g_stat": g_stat,
                    "g_crit": g_crit,
                    "pass": pass_no,
                    "direction": "spike" if candidate > mean else "drop",
                }
            )
            arr = np.delete(arr, max_pos)
            original_indices.pop(max_pos)
            pass_no += 1
        else:
            break

    return flagged


def detect_anomalies(df: pd.DataFrame, params: DashboardParams) -> pd.DataFrame:
    records: List[Dict[str, Any]] = []

    for mine_name, group in df.groupby("mine", sort=False):
        group = group.sort_values("date").copy().reset_index(drop=False)
        values = group["output"].astype(float)
        dates = group["date"]

        q1 = float(values.quantile(0.25))
        q3 = float(values.quantile(0.75))
        iqr = q3 - q1
        lower_iqr = q1 - params.iqr_k * iqr
        upper_iqr = q3 + params.iqr_k * iqr

        mean = float(values.mean())
        std = float(values.std(ddof=1)) if len(values) > 1 else 0.0
        if std > 0:
            z_scores = (values - mean) / std
        else:
            z_scores = pd.Series([0.0] * len(values), index=values.index)

        moving_avg = values.shift(1).rolling(window=params.ma_window, min_periods=params.ma_window).mean()
        with np.errstate(divide="ignore", invalid="ignore"):
            ma_distance_pct = np.where(
                moving_avg.abs() > 1e-12,
                np.abs((values - moving_avg) / moving_avg) * 100.0,
                np.nan,
            )

        grubbs_hits = {
            item["index"]: item
            for item in grubbs_iterative_flags(pd.Series(values.values, index=group["index"]), params.grubbs_alpha)
        }

        for i in range(len(group)):
            row_tests: List[str] = []
            details: Dict[str, Any] = {}
            value = float(values.iloc[i])

            if iqr > 0 and (value < lower_iqr or value > upper_iqr):
                row_tests.append("IQR rule")
                details["iqr"] = {
                    "lower_bound": lower_iqr,
                    "upper_bound": upper_iqr,
                    "q1": q1,
                    "q3": q3,
                    "iqr": iqr,
                    "k": params.iqr_k,
                }

            z_value = float(z_scores.iloc[i])
            if std > 0 and abs(z_value) > params.z_threshold:
                row_tests.append("z-score")
                details["z_score"] = {
                    "z": z_value,
                    "threshold": params.z_threshold,
                    "mean": mean,
                    "std": std,
                }

            ma_val = float(moving_avg.iloc[i]) if pd.notna(moving_avg.iloc[i]) else np.nan
            ma_pct = float(ma_distance_pct[i]) if not np.isnan(ma_distance_pct[i]) else np.nan
            if pd.notna(ma_val) and pd.notna(ma_pct) and ma_pct > params.ma_pct:
                row_tests.append("moving average distance")
                details["moving_average"] = {
                    "window": params.ma_window,
                    "moving_avg": ma_val,
                    "distance_pct": ma_pct,
                    "threshold_pct": params.ma_pct,
                }

            original_idx = int(group.loc[i, "index"])
            if original_idx in grubbs_hits:
                row_tests.append("Grubbs' test")
                details["grubbs"] = grubbs_hits[original_idx]

            if row_tests:
                if value > mean:
                    direction = "spike"
                elif value < mean:
                    direction = "drop"
                else:
                    direction = "anomaly"

                records.append(
                    {
                        "mine": mine_name,
                        "date": dates.iloc[i],
                        "value": value,
                        "direction": direction,
                        "tests": row_tests,
                        "details": details,
                    }
                )

    anomalies = pd.DataFrame(records)
    if anomalies.empty:
        return anomalies

    anomalies["date"] = pd.to_datetime(anomalies["date"]).dt.floor("D")
    anomalies = anomalies.sort_values(["mine", "date", "value"]).reset_index(drop=True)
    return anomalies


def build_series_payload(
    df: pd.DataFrame,
    anomalies: pd.DataFrame,
    entities: Optional[List[str]],
    trend_degree: int,
) -> List[Dict[str, Any]]:
    selected = entities or sorted(df["mine"].unique().tolist(), key=lambda x: (x != "Total", x))
    payload: List[Dict[str, Any]] = []

    anomaly_lookup: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    if not anomalies.empty:
        for _, row in anomalies.iterrows():
            key = (row["mine"], row["date"].date().isoformat())
            anomaly_lookup.setdefault(key, []).append(
                {
                    "tests": row["tests"],
                    "direction": row["direction"],
                    "value": round(float(row["value"]), 6),
                }
            )

    for mine_name in selected:
        group = df[df["mine"] == mine_name].sort_values("date").copy()
        if group.empty:
            continue

        x = group["date"].dt.strftime("%Y-%m-%d").tolist()
        y = group["output"].astype(float).round(6).tolist()

        n = len(group)
        degree = min(trend_degree, max(1, n - 1))
        x_num = np.arange(n)
        trend = [None] * n
        if n >= 2:
            coeffs = np.polyfit(x_num, group["output"].astype(float), degree)
            trend = np.polyval(coeffs, x_num).astype(float).round(6).tolist()

        outliers = []
        for _, row in group.iterrows():
            key = (mine_name, row["date"].date().isoformat())
            hits = anomaly_lookup.get(key, [])
            if hits:
                tests = sorted({test for hit in hits for test in hit["tests"]})
                direction = hits[0]["direction"]
                outliers.append(
                    {
                        "x": row["date"].strftime("%Y-%m-%d"),
                        "y": round(float(row["output"]), 6),
                        "tests": tests,
                        "direction": direction,
                    }
                )

        payload.append(
            {
                "mine": mine_name,
                "x": x,
                "y": y,
                "trendline": trend,
                "outliers": outliers,
            }
        )

    return payload


def enrich_metrics(metrics: pd.DataFrame, anomalies: pd.DataFrame) -> pd.DataFrame:
    metrics = metrics.copy()
    metrics["anomaly_count"] = 0
    metrics["spike_count"] = 0
    metrics["drop_count"] = 0
    metrics["iqr_count"] = 0
    metrics["z_score_count"] = 0
    metrics["moving_average_count"] = 0
    metrics["grubbs_count"] = 0

    if anomalies.empty:
        return metrics

    for idx, row in metrics.iterrows():
        mine_name = row["mine"]
        subset = anomalies[anomalies["mine"] == mine_name]
        metrics.at[idx, "anomaly_count"] = int(len(subset))
        metrics.at[idx, "spike_count"] = int((subset["direction"] == "spike").sum())
        metrics.at[idx, "drop_count"] = int((subset["direction"] == "drop").sum())

        def _count_test(test_name: str) -> int:
            return int(subset["tests"].apply(lambda x: test_name in x).sum())

        metrics.at[idx, "iqr_count"] = _count_test("IQR rule")
        metrics.at[idx, "z_score_count"] = _count_test("z-score")
        metrics.at[idx, "moving_average_count"] = _count_test("moving average distance")
        metrics.at[idx, "grubbs_count"] = _count_test("Grubbs' test")

    return metrics


def dataframe_to_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    for row in df.to_dict(orient="records"):
        clean_row = {}
        for key, value in row.items():
            if isinstance(value, pd.Timestamp):
                clean_row[key] = value.strftime("%Y-%m-%d")
            elif isinstance(value, (np.floating, float)):
                clean_row[key] = None if np.isnan(value) else round(float(value), 6)
            elif isinstance(value, (np.integer, int)):
                clean_row[key] = int(value)
            else:
                clean_row[key] = value
        result.append(clean_row)
    return result


def anomaly_records(anomalies: pd.DataFrame) -> List[Dict[str, Any]]:
    if anomalies.empty:
        return []

    records = []
    for _, row in anomalies.iterrows():
        records.append(
            {
                "mine": row["mine"],
                "date": row["date"].strftime("%Y-%m-%d"),
                "value": round(float(row["value"]), 6),
                "direction": row["direction"],
                "tests": row["tests"],
                "details": row["details"],
            }
        )
    return records


def build_dashboard_payload(params: DashboardParams) -> Dict[str, Any]:
    base_df = load_dataset(params.sheet_url)
    df = add_total_series(base_df)

    anomalies = detect_anomalies(df, params)
    metrics = compute_summary_stats(df)
    metrics = enrich_metrics(metrics, anomalies)

    if params.entities:
        selected_set = set(params.entities) | {"Total"}
        filtered_df = df[df["mine"].isin(selected_set)].copy()
        filtered_anomalies = anomalies[anomalies["mine"].isin(selected_set)].copy() if not anomalies.empty else anomalies
        filtered_metrics = metrics[metrics["mine"].isin(selected_set)].copy()
    else:
        filtered_df = df
        filtered_anomalies = anomalies
        filtered_metrics = metrics

    payload = {
        "dataset": {
            "sheet_url": normalize_google_sheet_url(params.sheet_url),
            "row_count": int(base_df.shape[0]),
            "min_date": base_df["date"].min().strftime("%Y-%m-%d"),
            "max_date": base_df["date"].max().strftime("%Y-%m-%d"),
            "mine_count": int(base_df["mine"].nunique()),
        },
        "params": {
            "chart_type": params.chart_type,
            "trend_degree": params.trend_degree,
            "iqr_k": params.iqr_k,
            "z_threshold": params.z_threshold,
            "ma_window": params.ma_window,
            "ma_pct": params.ma_pct,
            "grubbs_alpha": params.grubbs_alpha,
            "entities": params.entities,
        },
        "metrics": dataframe_to_records(filtered_metrics.sort_values(["mine"])),
        "anomalies": anomaly_records(filtered_anomalies),
        "chart": {
            "chart_type": params.chart_type,
            "series": build_series_payload(
                filtered_df,
                filtered_anomalies,
                params.entities,
                params.trend_degree,
            ),
        },
    }
    return payload


def _encode_fig_to_base64(fig: plt.Figure) -> str:
    buffer = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buffer, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def make_total_chart(df: pd.DataFrame, anomalies: pd.DataFrame, trend_degree: int) -> str:
    total = df[df["mine"] == "Total"].sort_values("date").copy()
    if total.empty:
        return ""

    dates = total["date"]
    values = total["output"].astype(float).to_numpy()

    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(dates, values, linewidth=2.2, label="Total output")

    if len(total) >= 2:
        x_num = np.arange(len(total))
        degree = min(trend_degree, max(1, len(total) - 1))
        coeffs = np.polyfit(x_num, values, degree)
        trend = np.polyval(coeffs, x_num)
        ax.plot(dates, trend, linestyle="--", linewidth=1.8, label=f"Polynomial trend (deg {degree})")

    if not anomalies.empty:
        out = anomalies[anomalies["mine"] == "Total"].copy()
        if not out.empty:
            merged = total.merge(out[["date", "direction"]], on="date", how="inner").drop_duplicates(subset=["date"])
            if not merged.empty:
                colors = np.where(merged["direction"] == "spike", "red", "orange")
                ax.scatter(merged["date"], merged["output"], c=colors, s=55, zorder=5, label="Outliers")

    ax.set_title("Total output with highlighted anomalies")
    ax.set_xlabel("Date")
    ax.set_ylabel("Output")
    ax.grid(alpha=0.25)
    ax.legend()
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
    fig.autofmt_xdate(rotation=35)
    return _encode_fig_to_base64(fig)


def make_mines_chart(df: pd.DataFrame, anomalies: pd.DataFrame, chart_type: str) -> str:
    mines_only = df[df["mine"] != "Total"].copy()
    if mines_only.empty:
        return ""

    pivot = mines_only.pivot(index="date", columns="mine", values="output").sort_index().fillna(0)

    fig, ax = plt.subplots(figsize=(11.5, 5.2))
    if chart_type == "bar":
        pivot.plot(kind="bar", ax=ax, width=0.85)
    elif chart_type == "stacked":
        pivot.plot(kind="bar", stacked=True, ax=ax, width=0.85)
    else:
        pivot.plot(kind="line", ax=ax, linewidth=2)

        if not anomalies.empty:
            out = anomalies[anomalies["mine"] != "Total"].copy()
            if not out.empty:
                for mine_name, group in out.groupby("mine"):
                    source = mines_only[mines_only["mine"] == mine_name]
                    merged = source.merge(group[["date", "direction"]], on="date", how="inner").drop_duplicates(subset=["date"])
                    if not merged.empty:
                        colors = np.where(merged["direction"] == "spike", "red", "orange")
                        ax.scatter(merged["date"], merged["output"], c=colors, s=40, zorder=5)

    ax.set_title(f"Mine output chart ({chart_type})")
    ax.set_xlabel("Date")
    ax.set_ylabel("Output")
    ax.grid(alpha=0.2)
    if chart_type == "line":
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))
        fig.autofmt_xdate(rotation=35)
    else:
        ax.tick_params(axis="x", rotation=40)
    return _encode_fig_to_base64(fig)


def build_report_html(payload: Dict[str, Any], total_chart_b64: str, mines_chart_b64: str) -> str:
    metrics_rows = []
    for row in payload["metrics"]:
        metrics_rows.append(
            f"""
            <tr>
                <td>{row['mine']}</td>
                <td>{row['mean_daily_output']:.2f}</td>
                <td>{row['std_dev']:.2f}</td>
                <td>{row['median']:.2f}</td>
                <td>{row['iqr']:.2f}</td>
                <td>{row['anomaly_count']}</td>
                <td>{row['spike_count']}</td>
                <td>{row['drop_count']}</td>
            </tr>
            """
        )

    spike_rows = []
    drop_rows = []
    for item in payload["anomalies"]:
        tests = ", ".join(item["tests"])
        html_row = f"""
        <tr>
            <td>{item['mine']}</td>
            <td>{item['date']}</td>
            <td>{item['value']:.2f}</td>
            <td>{tests}</td>
        </tr>
        """
        if item["direction"] == "spike":
            spike_rows.append(html_row)
        else:
            drop_rows.append(html_row)

    total_anomalies = len(payload["anomalies"])
    total_spikes = sum(1 for x in payload["anomalies"] if x["direction"] == "spike")
    total_drops = sum(1 for x in payload["anomalies"] if x["direction"] == "drop")

    return f"""
    <!doctype html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @page {{
            size: A4 landscape;
            margin: 18mm 14mm 16mm 14mm;
        }}
        body {{
            font-family: Arial, Helvetica, sans-serif;
            color: #1b263b;
            font-size: 12px;
        }}
        h1, h2, h3 {{
            margin: 0 0 10px 0;
            color: #102a43;
        }}
        h1 {{
            font-size: 24px;
        }}
        h2 {{
            font-size: 17px;
            margin-top: 20px;
            border-bottom: 2px solid #d9e2ec;
            padding-bottom: 6px;
        }}
        .muted {{
            color: #52606d;
        }}
        .grid {{
            display: table;
            width: 100%;
            border-collapse: separate;
            border-spacing: 10px;
            margin: 10px 0 12px 0;
        }}
        .card {{
            display: table-cell;
            width: 25%;
            background: #f7f9fc;
            border: 1px solid #d9e2ec;
            border-radius: 10px;
            padding: 12px;
        }}
        .card .label {{
            font-size: 11px;
            color: #52606d;
        }}
        .card .value {{
            font-size: 20px;
            font-weight: bold;
            margin-top: 6px;
            color: #102a43;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
        }}
        th, td {{
            border: 1px solid #d9e2ec;
            padding: 7px 8px;
            vertical-align: top;
        }}
        th {{
            background: #eef2f7;
            text-align: left;
        }}
        .chart {{
            margin-top: 10px;
            text-align: center;
        }}
        .chart img {{
            width: 100%;
            max-height: 420px;
            object-fit: contain;
            border: 1px solid #d9e2ec;
            border-radius: 8px;
        }}
        .note {{
            background: #fffbea;
            border: 1px solid #ffe08a;
            padding: 10px 12px;
            border-radius: 8px;
            margin-top: 10px;
        }}
        .empty {{
            color: #7b8794;
            font-style: italic;
            padding: 10px 0;
        }}
      </style>
    </head>
    <body>
      <h1>Mining Output Analysis Report</h1>
      <div class="muted">
        Source: {payload['dataset']['sheet_url']}<br>
        Coverage: {payload['dataset']['min_date']} to {payload['dataset']['max_date']}<br>
        Mines: {payload['dataset']['mine_count']} | Raw rows: {payload['dataset']['row_count']}
      </div>

      <div class="grid">
        <div class="card">
          <div class="label">Total anomalies</div>
          <div class="value">{total_anomalies}</div>
        </div>
        <div class="card">
          <div class="label">Spike anomalies</div>
          <div class="value">{total_spikes}</div>
        </div>
        <div class="card">
          <div class="label">Drop anomalies</div>
          <div class="value">{total_drops}</div>
        </div>
        <div class="card">
          <div class="label">Trend degree / Chart type</div>
          <div class="value">{payload['params']['trend_degree']} / {payload['params']['chart_type']}</div>
        </div>
      </div>

      <div class="note">
        Detection settings: IQR k = {payload['params']['iqr_k']}, z-threshold = {payload['params']['z_threshold']},
        moving average window = {payload['params']['ma_window']},
        moving average distance = {payload['params']['ma_pct']}%,
        Grubbs alpha = {payload['params']['grubbs_alpha']}.
      </div>

      <h2>Summary Statistics</h2>
      <table>
        <thead>
          <tr>
            <th>Mine</th>
            <th>Mean daily output</th>
            <th>Standard deviation</th>
            <th>Median</th>
            <th>IQR</th>
            <th>Anomalies</th>
            <th>Spikes</th>
            <th>Drops</th>
          </tr>
        </thead>
        <tbody>
          {''.join(metrics_rows)}
        </tbody>
      </table>

      <h2>Total Output Chart</h2>
      <div class="chart">
        <img src="data:image/png;base64,{total_chart_b64}">
      </div>

      <h2>Mine Output Chart</h2>
      <div class="chart">
        <img src="data:image/png;base64,{mines_chart_b64}">
      </div>

      <h2>Spike Anomalies</h2>
      {
        '<table><thead><tr><th>Mine</th><th>Date</th><th>Value</th><th>Triggered tests</th></tr></thead><tbody>' + ''.join(spike_rows) + '</tbody></table>'
        if spike_rows else '<div class="empty">No spike anomalies detected for the selected parameters.</div>'
      }

      <h2>Drop Anomalies</h2>
      {
        '<table><thead><tr><th>Mine</th><th>Date</th><th>Value</th><th>Triggered tests</th></tr></thead><tbody>' + ''.join(drop_rows) + '</tbody></table>'
        if drop_rows else '<div class="empty">No drop anomalies detected for the selected parameters.</div>'
      }
    </body>
    </html>
    """


@app.route("/")
def index():
    return render_template("dashboard.html", default_sheet_url=DEFAULT_SHEET_URL)


@app.route("/api/dashboard")
def dashboard_api():
    try:
        params = parse_params()
        payload = build_dashboard_payload(params)
        return jsonify(payload)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.route("/report.pdf")
def report_pdf():
    if HTML is None:
        return jsonify(
            {
                "error": "WeasyPrint is not installed. Install it to enable PDF export."
            }
        ), 500

    try:
        params = parse_params()
        payload = build_dashboard_payload(params)

        base_df = load_dataset(params.sheet_url)
        df = add_total_series(base_df)
        anomalies = detect_anomalies(df, params)

        if params.entities:
            selected_set = set(params.entities) | {"Total"}
            df = df[df["mine"].isin(selected_set)].copy()
            anomalies = anomalies[anomalies["mine"].isin(selected_set)].copy() if not anomalies.empty else anomalies

        total_chart_b64 = make_total_chart(df, anomalies, params.trend_degree)
        mines_chart_b64 = make_mines_chart(df, anomalies, params.chart_type)
        html = build_report_html(payload, total_chart_b64, mines_chart_b64)

        pdf_io = io.BytesIO()
        HTML(string=html).write_pdf(target=pdf_io)
        pdf_io.seek(0)

        return send_file(
            pdf_io,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="mining_output_analysis_report.pdf",
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


if __name__ == "__main__":
    app.run(debug=False)
