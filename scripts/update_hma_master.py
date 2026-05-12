from pathlib import Path
import json
import math
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule


BASE_DIR = Path(__file__).resolve().parent.parent
DOWNLOADS_DIR = BASE_DIR / "downloads"
HISTORY_DIR = BASE_DIR / "historico"
MASTER_FILE = HISTORY_DIR / "HMA_Master.xlsx"

HISTORY_DIR.mkdir(exist_ok=True)

# Criterio comercial base para presupuestos diarios altos.
# Ajustable más adelante por cliente.
CHANGE_NOISE_PCT = 2.0
CHANGE_MODERATE_PCT = 5.0
CHANGE_CRITICAL_PCT = 10.0

DAILY_BUDGET_USD = 30000.0
HOURLY_BUDGET_USD = DAILY_BUDGET_USD / 24

MAX_SCALE_UP_PCT = 10.0
MAX_SCALE_DOWN_PCT = 10.0


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None:
            return default
        if isinstance(value, float) and math.isnan(value):
            return default
        return float(value)
    except Exception:
        return default


def safe_int(value: Any, default: int = 0) -> int:
    try:
        if value is None:
            return default
        if isinstance(value, float) and math.isnan(value):
            return default
        return int(float(value))
    except Exception:
        return default


def pct_change(current: float, previous: float) -> float:
    if previous == 0:
        return 0.0
    return ((current - previous) / previous) * 100


def direction(change_pct: float) -> str:
    if abs(change_pct) < CHANGE_NOISE_PCT:
        return "estable"
    if change_pct > 0:
        return "subió"
    return "bajó"


def relevance(change_pct: float) -> str:
    abs_change = abs(change_pct)

    if abs_change < CHANGE_NOISE_PCT:
        return "ruido normal"
    if abs_change < CHANGE_MODERATE_PCT:
        return "leve"
    if abs_change < CHANGE_CRITICAL_PCT:
        return "moderada"
    return "crítica"


def metric_business_interpretation(metric: str, change_pct: float) -> str:
    metric = metric.lower()

    if abs(change_pct) < CHANGE_NOISE_PCT:
        return "Variación menor al umbral. No tomar decisión por esta señal aislada."

    if metric == "cpa":
        return (
            "CPA subió: eficiencia deteriorada."
            if change_pct > 0
            else "CPA bajó: eficiencia mejorada."
        )

    if metric == "roas":
        return (
            "ROAS subió: rentabilidad mejorada."
            if change_pct > 0
            else "ROAS bajó: rentabilidad deteriorada."
        )

    if metric == "ctr":
        return (
            "CTR subió: mejor respuesta creativa o mayor relevancia."
            if change_pct > 0
            else "CTR bajó: posible problema de títulos, descripciones, imágenes o relevancia."
        )

    if metric == "cvr":
        return (
            "CVR subió: mejor calidad de tráfico o conversión."
            if change_pct > 0
            else "CVR bajó: posible problema de intención, landing, oferta o tracking."
        )

    if metric == "spend":
        return (
            "Gasto subió: revisar si el incremento está acompañado por conversiones."
            if change_pct > 0
            else "Gasto bajó: revisar pacing y entrega."
        )

    if metric == "clicks":
        return (
            "Clics subieron: revisar si las conversiones acompañan."
            if change_pct > 0
            else "Clics bajaron: revisar tráfico, anuncios o pérdida de volumen."
        )

    if metric == "conversions":
        return (
            "Conversiones subieron: señal positiva si CPA/ROAS acompañan."
            if change_pct > 0
            else "Conversiones bajaron: revisar tráfico, funnel o tracking."
        )

    return "Cambio registrado."


def find_artifact_runs() -> list[dict]:
    """
    Busca artifacts descargados en downloads/ y detecta pares:
    latest_summary.json + latest_metrics.csv + connection_status.json.
    """
    runs = []

    if not DOWNLOADS_DIR.exists():
        return runs

    summary_files = list(DOWNLOADS_DIR.rglob("latest_summary.json"))

    for summary_path in summary_files:
        exports_dir = summary_path.parent
        metrics_path = exports_dir / "latest_metrics.csv"
        connection_path = exports_dir / "connection_status.json"

        if not metrics_path.exists():
            continue

        artifact_root = exports_dir.parent
        reports_dir = artifact_root / "reports"

        named_report = ""
        latest_report = ""

        if reports_dir.exists():
            latest_candidate = reports_dir / "latest_hourly_report.md"
            if latest_candidate.exists():
                latest_report = str(latest_candidate)

            named_reports = [
                p for p in reports_dir.glob("*.md")
                if p.name != "latest_hourly_report.md"
            ]
            if named_reports:
                named_report = str(named_reports[0])

        run_key = str(artifact_root.relative_to(DOWNLOADS_DIR))

        runs.append(
            {
                "run_key": run_key,
                "artifact_root": artifact_root,
                "summary_path": summary_path,
                "metrics_path": metrics_path,
                "connection_path": connection_path if connection_path.exists() else None,
                "latest_report": latest_report,
                "named_report": named_report,
            }
        )

    return runs


def load_summary(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def build_hourly_summary(runs: list[dict]) -> pd.DataFrame:
    rows = []

    for run in runs:
        summary = load_summary(run["summary_path"])
        latest = summary.get("latest", {})
        thresholds = summary.get("thresholds", {})
        changes = summary.get("changes", {})
        alerts = summary.get("alerts", [])
        connection_status = summary.get("connection_status", {})

        rows.append(
            {
                "run_key": run["run_key"],
                "client_number": summary.get("client_number", ""),
                "timestamp": summary.get("latest_timestamp", ""),
                "timezone": summary.get("report_timezone", ""),
                "data_source": summary.get("source", ""),
                "requested_data_source": summary.get("requested_data_source", ""),
                "health_status": summary.get("health_status", ""),
                "spend": safe_float(latest.get("spend")),
                "impressions": safe_int(latest.get("impressions")),
                "clicks": safe_int(latest.get("clicks")),
                "ctr": safe_float(latest.get("ctr")),
                "cpc": safe_float(latest.get("cpc")),
                "conversions": safe_int(latest.get("conversions")),
                "cvr": safe_float(latest.get("cvr")),
                "cpa": safe_float(latest.get("cpa")),
                "revenue": safe_float(latest.get("revenue")),
                "roas": safe_float(latest.get("roas")),
                "alerts_count": len(alerts),
                "cpa_threshold": safe_float(thresholds.get("cpa_threshold")),
                "roas_threshold": safe_float(thresholds.get("roas_threshold")),
                "ctr_threshold": safe_float(thresholds.get("ctr_threshold")),
                "spend_change_pct_report": safe_float(changes.get("spend_change_pct")),
                "cpa_change_pct_report": safe_float(changes.get("cpa_change_pct")),
                "roas_change_pct_report": safe_float(changes.get("roas_change_pct")),
                "google_ads_ready": connection_status.get("google_ads_ready", False),
                "meta_ads_ready": connection_status.get("meta_ads_ready", False),
                "named_report_path": run["named_report"],
                "latest_report_path": run["latest_report"],
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df["timestamp_dt"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.sort_values("timestamp_dt").drop_duplicates(subset=["run_key"], keep="last")

    return df


def build_campaign_metrics(runs: list[dict]) -> pd.DataFrame:
    frames = []

    for run in runs:
        try:
            summary = load_summary(run["summary_path"])
            latest_timestamp = summary.get("latest_timestamp", "")
            client_number = summary.get("client_number", "")

            df = pd.read_csv(run["metrics_path"])
            df["run_key"] = run["run_key"]
            df["client_number"] = client_number
            df["summary_timestamp"] = latest_timestamp

            frames.append(df)
        except Exception:
            continue

    if not frames:
        return pd.DataFrame()

    df_all = pd.concat(frames, ignore_index=True)
    df_all["summary_timestamp_dt"] = pd.to_datetime(
        df_all["summary_timestamp"],
        errors="coerce",
    )

    return df_all.sort_values(["summary_timestamp_dt", "platform", "campaign_name"])


def build_metric_comparison(hourly_df: pd.DataFrame) -> pd.DataFrame:
    if hourly_df.empty or len(hourly_df) < 2:
        return pd.DataFrame()

    hourly_df = hourly_df.sort_values("timestamp_dt").reset_index(drop=True)

    metrics = [
        "spend",
        "impressions",
        "clicks",
        "ctr",
        "cpc",
        "conversions",
        "cvr",
        "cpa",
        "revenue",
        "roas",
    ]

    rows = []

    for idx in range(1, len(hourly_df)):
        current = hourly_df.iloc[idx]
        previous = hourly_df.iloc[idx - 1]

        for metric in metrics:
            current_value = safe_float(current.get(metric))
            previous_value = safe_float(previous.get(metric))
            change = pct_change(current_value, previous_value)

            rows.append(
                {
                    "timestamp": current["timestamp"],
                    "client_number": current["client_number"],
                    "metric": metric,
                    "current_value": current_value,
                    "previous_value": previous_value,
                    "change_pct": change,
                    "direction": direction(change),
                    "relevance": relevance(change),
                    "commercial_interpretation": metric_business_interpretation(
                        metric,
                        change,
                    ),
                }
            )

    return pd.DataFrame(rows)


def get_last_change(comparison_df: pd.DataFrame, metric: str) -> float:
    if comparison_df.empty:
        return 0.0

    metric_df = comparison_df[comparison_df["metric"] == metric]
    if metric_df.empty:
        return 0.0

    return safe_float(metric_df.iloc[-1]["change_pct"])


def confidence_from_signals(
    negative_signals: int,
    positive_signals: int,
    history_len: int,
) -> str:
    if history_len < 2:
        return "BAJA"

    if negative_signals >= 3 or positive_signals >= 3:
        return "ALTA" if history_len >= 3 else "MEDIA"

    if negative_signals >= 2 or positive_signals >= 2:
        return "MEDIA"

    return "BAJA"


def build_recommendations(hourly_df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
    if hourly_df.empty:
        return pd.DataFrame()

    latest = hourly_df.sort_values("timestamp_dt").iloc[-1]
    history_len = len(hourly_df)

    cpa_change = get_last_change(comparison_df, "cpa")
    roas_change = get_last_change(comparison_df, "roas")
    ctr_change = get_last_change(comparison_df, "ctr")
    cvr_change = get_last_change(comparison_df, "cvr")
    spend_change = get_last_change(comparison_df, "spend")
    clicks_change = get_last_change(comparison_df, "clicks")
    conversions_change = get_last_change(comparison_df, "conversions")

    cpa = safe_float(latest.get("cpa"))
    roas = safe_float(latest.get("roas"))
    ctr = safe_float(latest.get("ctr"))
    conversions = safe_int(latest.get("conversions"))
    clicks = safe_int(latest.get("clicks"))
    spend = safe_float(latest.get("spend"))

    cpa_threshold = safe_float(latest.get("cpa_threshold"), 25)
    roas_threshold = safe_float(latest.get("roas_threshold"), 1.5)
    ctr_threshold = safe_float(latest.get("ctr_threshold"), 1.0)

    rows = []

    def add(
        level: str,
        action: str,
        motive: str,
        area: str,
        confidence: str,
        human_review: bool = True,
    ):
        rows.append(
            {
                "timestamp": latest.get("timestamp", ""),
                "client_number": latest.get("client_number", ""),
                "level": level,
                "area": area,
                "recommended_action": action,
                "motive": motive,
                "confidence": confidence,
                "requires_human_review": human_review,
                "daily_budget_usd_reference": DAILY_BUDGET_USD,
                "hourly_budget_usd_reference": HOURLY_BUDGET_USD,
            }
        )

    negative_signals = 0
    positive_signals = 0

    if cpa_change > CHANGE_MODERATE_PCT:
        negative_signals += 1
    if roas_change < -CHANGE_MODERATE_PCT:
        negative_signals += 1
    if ctr_change < -CHANGE_MODERATE_PCT:
        negative_signals += 1
    if cvr_change < -CHANGE_MODERATE_PCT:
        negative_signals += 1
    if conversions_change < -CHANGE_MODERATE_PCT:
        negative_signals += 1

    if cpa_change < -CHANGE_MODERATE_PCT:
        positive_signals += 1
    if roas_change > CHANGE_MODERATE_PCT:
        positive_signals += 1
    if conversions_change > CHANGE_MODERATE_PCT:
        positive_signals += 1
    if cvr_change > CHANGE_MODERATE_PCT:
        positive_signals += 1

    confidence = confidence_from_signals(
        negative_signals=negative_signals,
        positive_signals=positive_signals,
        history_len=history_len,
    )

    if spend > 0 and clicks > 100 and conversions == 0:
        add(
            "CRÍTICO",
            "No cambiar presupuesto hasta validar tracking. Revisar pixel, conversion API, tags, evento de conversión, landing y caída del sitio.",
            "Hay gasto y volumen de clics, pero conversiones en cero.",
            "tracking/funnel",
            "MEDIA" if history_len >= 2 else "BAJA",
        )

    if cpa > cpa_threshold and cpa_change > CHANGE_CRITICAL_PCT and roas_change < -CHANGE_MODERATE_PCT:
        add(
            "ALTO",
            f"Frenar escalado o reducir presupuesto hasta {MAX_SCALE_DOWN_PCT:.0f}% con revisión humana.",
            f"CPA por encima del umbral y subiendo {cpa_change:.1f}%, con ROAS cayendo {roas_change:.1f}%.",
            "presupuesto/eficiencia",
            confidence,
        )

    elif cpa > cpa_threshold and cpa_change > CHANGE_MODERATE_PCT:
        add(
            "MEDIO",
            "Mantener presupuesto y revisar eficiencia antes de escalar.",
            f"CPA supera umbral y subió {cpa_change:.1f}%.",
            "presupuesto/eficiencia",
            confidence,
        )

    if roas < roas_threshold and roas_change < -CHANGE_MODERATE_PCT:
        add(
            "ALTO" if abs(roas_change) >= CHANGE_CRITICAL_PCT else "MEDIO",
            "Revisar distribución de presupuesto y calidad del tráfico. No aumentar inversión hasta recuperar ROAS.",
            f"ROAS bajo umbral y variación de {roas_change:.1f}%.",
            "rentabilidad",
            confidence,
        )

    if ctr < ctr_threshold or ctr_change < -CHANGE_MODERATE_PCT:
        add(
            "MEDIO" if abs(ctr_change) < CHANGE_CRITICAL_PCT else "ALTO",
            "Revisar títulos, descripciones, imágenes, fatiga creativa y relevancia del anuncio.",
            f"CTR actual {ctr:.2f}% y cambio {ctr_change:.1f}%.",
            "creatividad",
            confidence,
        )

    if clicks_change > CHANGE_MODERATE_PCT and conversions_change <= 0 and spend_change > 0:
        add(
            "ALTO" if clicks_change >= CHANGE_CRITICAL_PCT else "MEDIO",
            "Controlar keywords, términos de búsqueda, negativas, concordancias y calidad del tráfico.",
            f"Clics suben {clicks_change:.1f}%, gasto sube {spend_change:.1f}%, pero conversiones no acompañan.",
            "keywords/tráfico",
            confidence,
        )

    if cvr_change < -CHANGE_MODERATE_PCT and ctr_change >= -CHANGE_NOISE_PCT:
        add(
            "MEDIO",
            "Revisar landing, oferta, velocidad, formularios y evento de conversión.",
            f"CVR cae {cvr_change:.1f}% mientras el CTR no cae de forma equivalente.",
            "funnel/landing",
            confidence,
        )

    if (
        cpa_change < -CHANGE_MODERATE_PCT
        and roas_change > CHANGE_MODERATE_PCT
        and conversions_change > 0
        and cvr_change >= -CHANGE_NOISE_PCT
    ):
        add(
            "OPORTUNIDAD",
            f"Considerar aumento gradual de presupuesto de 5% a {MAX_SCALE_UP_PCT:.0f}% con revisión humana.",
            f"CPA baja {abs(cpa_change):.1f}%, ROAS sube {roas_change:.1f}% y conversiones acompañan.",
            "escalado",
            confidence,
        )

    if not rows:
        add(
            "NORMAL",
            "Mantener presupuesto y continuar monitoreo horario.",
            "No hay deterioro relevante por encima de los umbrales comerciales configurados.",
            "monitoreo",
            "BAJA" if history_len < 3 else "MEDIA",
            human_review=False,
        )

    return pd.DataFrame(rows)


def build_creative_assets(campaign_df: pd.DataFrame) -> pd.DataFrame:
    if campaign_df.empty:
        return pd.DataFrame()

    df = campaign_df.copy()

    expected_cols = [
        "ad_title",
        "ad_description",
        "image_asset_url",
        "image_file",
        "creative_group",
    ]

    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    def creative_recommendation(row):
        ctr = safe_float(row.get("ctr"))
        cvr = safe_float(row.get("cvr"))
        cpa = safe_float(row.get("cpa"))
        roas = safe_float(row.get("roas"))

        if ctr < 1.0:
            return "Revisar títulos, descripciones e imágenes. Posible baja relevancia o fatiga creativa."

        if ctr >= 1.0 and cvr < 1.0:
            return "El anuncio consigue clics, pero la conversión es débil. Revisar intención, landing u oferta."

        if cpa > 25 and roas < 1.5:
            return "Creatividad y tráfico no están generando eficiencia suficiente. Revisar ángulo, promesa y segmentación."

        return "Mantener creatividad bajo monitoreo."

    df["creative_recommendation"] = df.apply(creative_recommendation, axis=1)

    wanted_cols = [
        "summary_timestamp",
        "client_number",
        "platform",
        "account_name",
        "campaign_name",
        "creative_group",
        "ad_title",
        "ad_description",
        "image_asset_url",
        "image_file",
        "spend",
        "impressions",
        "clicks",
        "ctr",
        "conversions",
        "cvr",
        "cpa",
        "roas",
        "creative_recommendation",
    ]

    for col in wanted_cols:
        if col not in df.columns:
            df[col] = ""

    return df[wanted_cols]


def autosize_and_style(path: Path):
    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if ws.max_row >= 1:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        for column_cells in ws.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))

            width = min(max(max_length + 2, 10), 45)
            ws.column_dimensions[column_letter].width = width

    if "metric_comparison" in wb.sheetnames:
        ws = wb["metric_comparison"]
        ws.conditional_formatting.add(
            "F2:F10000",
            CellIsRule(
                operator="greaterThan",
                formula=["10"],
                fill=PatternFill("solid", fgColor="FFC7CE"),
            ),
        )
        ws.conditional_formatting.add(
            "F2:F10000",
            CellIsRule(
                operator="lessThan",
                formula=["-10"],
                fill=PatternFill("solid", fgColor="C6EFCE"),
            ),
        )

    wb.save(path)


def main():
    runs = find_artifact_runs()

    hourly_df = build_hourly_summary(runs)
    campaign_df = build_campaign_metrics(runs)
    comparison_df = build_metric_comparison(hourly_df)
    recommendations_df = build_recommendations(hourly_df, comparison_df)
    creative_df = build_creative_assets(campaign_df)

    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl") as writer:
        hourly_df.to_excel(writer, sheet_name="hourly_summary", index=False)
        campaign_df.to_excel(writer, sheet_name="campaign_metrics", index=False)
        comparison_df.to_excel(writer, sheet_name="metric_comparison", index=False)
        recommendations_df.to_excel(writer, sheet_name="recommendations", index=False)
        creative_df.to_excel(writer, sheet_name="creative_assets", index=False)

    autosize_and_style(MASTER_FILE)

    print(f"HMA Master actualizado correctamente: {MASTER_FILE}")
    print(f"Runs procesados: {len(runs)}")
    print(f"Filas hourly_summary: {len(hourly_df)}")
    print(f"Filas campaign_metrics: {len(campaign_df)}")
    print(f"Filas metric_comparison: {len(comparison_df)}")
    print(f"Filas recommendations: {len(recommendations_df)}")
    print(f"Filas creative_assets: {len(creative_df)}")


if __name__ == "__main__":
    main()
