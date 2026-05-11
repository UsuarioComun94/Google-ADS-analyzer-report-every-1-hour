from pathlib import Path
import datetime as dt
from zoneinfo import ZoneInfo
import random
import os
import logging
import json

import pandas as pd
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "reports"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"

REPORTS_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)
EXPORTS_DIR.mkdir(exist_ok=True)

load_dotenv(BASE_DIR / ".env")

DATA_SOURCE = os.getenv("DATA_SOURCE", "simulated").strip().lower()

CPA_THRESHOLD = float(os.getenv("CPA_THRESHOLD", 25))
ROAS_THRESHOLD = float(os.getenv("ROAS_THRESHOLD", 1.5))
CTR_THRESHOLD = float(os.getenv("CTR_THRESHOLD", 1.0))

CLIENT_NUMBER = os.getenv("CLIENT_NUMBER", "CLIENTE-DEMO-0001")
REPORT_SUFFIX = os.getenv("REPORT_SUFFIX", "JPPQ")
REPORT_TIMEZONE = os.getenv("REPORT_TIMEZONE", "America/Argentina/Cordoba")

GOOGLE_ADS_REQUIRED_ENV = [
    "GOOGLE_ADS_CLIENT_ID",
    "GOOGLE_ADS_CLIENT_SECRET",
    "GOOGLE_ADS_DEVELOPER_TOKEN",
    "GOOGLE_ADS_REFRESH_TOKEN",
    "GOOGLE_ADS_CUSTOMER_ID",
]

META_ADS_REQUIRED_ENV = [
    "META_ACCESS_TOKEN",
    "META_AD_ACCOUNT_ID",
]


logging.basicConfig(
    filename=LOGS_DIR / "hma.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


def get_local_now() -> dt.datetime:
    """Devuelve fecha/hora en la zona horaria configurada."""
    return dt.datetime.now(ZoneInfo(REPORT_TIMEZONE))


def safe_div(numerator: float, denominator: float) -> float:
    """Evita división por cero."""
    return numerator / denominator if denominator else 0


def sanitize_filename_part(value: str) -> str:
    """Convierte un texto en una parte segura para nombre de archivo."""
    return (
        str(value)
        .strip()
        .replace(" ", "-")
        .replace("/", "-")
        .replace("\\", "-")
        .replace(":", "-")
        .replace('"', "")
        .replace("'", "")
        .replace("ñ", "n")
        .replace("Ñ", "N")
    )


def build_report_basename(run_time: dt.datetime) -> str:
    """Construye nombre seguro para archivo y artifact."""
    safe_client = sanitize_filename_part(CLIENT_NUMBER)
    safe_suffix = sanitize_filename_part(REPORT_SUFFIX)
    timestamp = run_time.strftime("%Y-%m-%d_%H-%M")

    return f"Campania_{safe_client}_{timestamp}_{safe_suffix}"


def missing_env_vars(required_vars: list[str]) -> list[str]:
    """Devuelve variables requeridas que no están configuradas."""
    return [var for var in required_vars if not os.getenv(var)]


def build_connection_status() -> dict:
    """Construye estado de conexión para APIs reales."""
    missing_google = missing_env_vars(GOOGLE_ADS_REQUIRED_ENV)
    missing_meta = missing_env_vars(META_ADS_REQUIRED_ENV)

    google_ready = len(missing_google) == 0
    meta_ready = len(missing_meta) == 0

    real_api_execution_enabled = False

    notes = []

    if DATA_SOURCE == "simulated":
        notes.append(
            "El sistema está ejecutándose en modo simulado. No usa datos reales de APIs."
        )
    elif DATA_SOURCE == "google_ads":
        notes.append(
            "DATA_SOURCE solicita Google Ads, pero el extractor real todavía no está implementado."
        )
    elif DATA_SOURCE == "meta_ads":
        notes.append(
            "DATA_SOURCE solicita Meta Ads, pero el extractor real todavía no está implementado."
        )
    else:
        notes.append(
            f"DATA_SOURCE='{DATA_SOURCE}' no es un modo reconocido. Se mantiene ejecución simulada."
        )

    if not google_ready:
        notes.append("Faltan credenciales para Google Ads API.")

    if not meta_ready:
        notes.append("Faltan credenciales para Meta Marketing API.")

    return {
        "requested_data_source": DATA_SOURCE,
        "effective_data_source": "simulated_data",
        "real_api_execution_enabled": real_api_execution_enabled,
        "google_ads_ready": google_ready,
        "meta_ads_ready": meta_ready,
        "missing_google_ads_credentials": missing_google,
        "missing_meta_ads_credentials": missing_meta,
        "configured_google_ads_credentials": [
            var for var in GOOGLE_ADS_REQUIRED_ENV if var not in missing_google
        ],
        "configured_meta_ads_credentials": [
            var for var in META_ADS_REQUIRED_ENV if var not in missing_meta
        ],
        "notes": notes,
    }


def simulate_ads_data() -> pd.DataFrame:
    """
    Simula datos horarios de campañas.

    Esta función representa el módulo extractor.
    Más adelante se reemplaza por:
    - fetch_google_ads_data()
    - fetch_meta_ads_data()
    """
    now = get_local_now().replace(minute=0, second=0, microsecond=0)

    campaigns = [
        {
            "platform": "google_ads",
            "account_name": "Demo Account",
            "campaign_name": "Search - High Intent Leads",
        },
        {
            "platform": "google_ads",
            "account_name": "Demo Account",
            "campaign_name": "Performance Max - Lead Gen",
        },
        {
            "platform": "meta_ads",
            "account_name": "Demo Account",
            "campaign_name": "Meta - Cold Prospecting",
        },
        {
            "platform": "meta_ads",
            "account_name": "Demo Account",
            "campaign_name": "Meta - Retargeting",
        },
    ]

    rows = []

    for hour_offset in range(24):
        timestamp = now - dt.timedelta(hours=hour_offset)

        for campaign in campaigns:
            impressions = random.randint(900, 18000)
            clicks = random.randint(20, 650)
            conversions = random.randint(0, 22)
            spend = round(random.uniform(15, 220), 2)

            if conversions > 0:
                revenue = round(conversions * random.uniform(35, 160), 2)
            else:
                revenue = 0.0

            rows.append(
                {
                    "timestamp": timestamp,
                    "platform": campaign["platform"],
                    "account_name": campaign["account_name"],
                    "campaign_name": campaign["campaign_name"],
                    "spend": spend,
                    "impressions": impressions,
                    "clicks": clicks,
                    "conversions": conversions,
                    "revenue": revenue,
                }
            )

    df = pd.DataFrame(rows)
    return df.sort_values("timestamp", ascending=False).reset_index(drop=True)


def calculate_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """Calcula KPIs derivados."""
    df = df.copy()

    df["ctr"] = df.apply(
        lambda row: safe_div(row["clicks"], row["impressions"]) * 100,
        axis=1,
    )

    df["cpc"] = df.apply(
        lambda row: safe_div(row["spend"], row["clicks"]),
        axis=1,
    )

    df["cvr"] = df.apply(
        lambda row: safe_div(row["conversions"], row["clicks"]) * 100,
        axis=1,
    )

    df["cpa"] = df.apply(
        lambda row: safe_div(row["spend"], row["conversions"]),
        axis=1,
    )

    df["roas"] = df.apply(
        lambda row: safe_div(row["revenue"], row["spend"]),
        axis=1,
    )

    return df


def summarize_hour(df: pd.DataFrame, timestamp) -> dict:
    """Agrupa todas las campañas de una hora."""
    hour_df = df[df["timestamp"] == timestamp]

    summary = {
        "timestamp": timestamp,
        "spend": float(hour_df["spend"].sum()),
        "impressions": int(hour_df["impressions"].sum()),
        "clicks": int(hour_df["clicks"].sum()),
        "conversions": int(hour_df["conversions"].sum()),
        "revenue": float(hour_df["revenue"].sum()),
    }

    summary["ctr"] = safe_div(summary["clicks"], summary["impressions"]) * 100
    summary["cpc"] = safe_div(summary["spend"], summary["clicks"])
    summary["cvr"] = safe_div(summary["conversions"], summary["clicks"]) * 100
    summary["cpa"] = safe_div(summary["spend"], summary["conversions"])
    summary["roas"] = safe_div(summary["revenue"], summary["spend"])

    return summary


def percentage_change(current: float, previous: float) -> float:
    """Calcula variación porcentual."""
    if previous == 0:
        return 0
    return ((current - previous) / previous) * 100


def generate_alerts(latest: dict) -> list[str]:
    """Genera alertas operativas básicas."""
    alerts = []

    if latest["spend"] > 0 and latest["conversions"] == 0:
        alerts.append("⚠️ Gasto activo sin conversiones en la última hora.")

    if latest["cpa"] > CPA_THRESHOLD:
        alerts.append(
            f"⚠️ CPA alto: ${latest['cpa']:.2f}. Umbral configurado: ${CPA_THRESHOLD:.2f}."
        )

    if latest["roas"] < ROAS_THRESHOLD:
        alerts.append(
            f"⚠️ ROAS bajo: {latest['roas']:.2f}. Umbral configurado: {ROAS_THRESHOLD:.2f}."
        )

    if latest["ctr"] < CTR_THRESHOLD:
        alerts.append(
            f"⚠️ CTR bajo: {latest['ctr']:.2f}%. Umbral configurado: {CTR_THRESHOLD:.2f}%."
        )

    if latest["clicks"] > 100 and latest["conversions"] == 0:
        alerts.append(
            "⚠️ Volumen de clics relevante sin conversiones. Revisar tracking, oferta o calidad del tráfico."
        )

    if not alerts:
        alerts.append("✅ Rendimiento dentro de parámetros normales.")

    return alerts


def classify_health(alerts: list[str]) -> str:
    """Clasifica estado general del sistema."""
    critical_alerts = [alert for alert in alerts if "⚠️" in alert]

    if len(critical_alerts) >= 3:
        return "CRÍTICO"

    if len(critical_alerts) >= 1:
        return "ATENCIÓN"

    return "NORMAL"


def get_latest_campaigns(df: pd.DataFrame, latest_timestamp) -> pd.DataFrame:
    """Devuelve el detalle por campaña de la última hora."""
    latest_campaigns = df[df["timestamp"] == latest_timestamp].copy()
    latest_campaigns = latest_campaigns.sort_values("spend", ascending=False)

    return latest_campaigns


def generate_markdown_report(
    df: pd.DataFrame,
    connection_status: dict,
) -> tuple[str, dict, pd.DataFrame]:
    """Genera reporte Markdown y devuelve también summary + campañas."""
    timestamps = sorted(df["timestamp"].unique(), reverse=True)

    latest_timestamp = timestamps[0]
    previous_timestamp = timestamps[1]

    latest = summarize_hour(df, latest_timestamp)
    previous = summarize_hour(df, previous_timestamp)

    alerts = generate_alerts(latest)
    health_status = classify_health(alerts)

    spend_change = percentage_change(latest["spend"], previous["spend"])
    clicks_change = percentage_change(latest["clicks"], previous["clicks"])
    conversions_change = latest["conversions"] - previous["conversions"]
    cpa_change = percentage_change(latest["cpa"], previous["cpa"])
    roas_change = percentage_change(latest["roas"], previous["roas"])

    latest_campaigns = get_latest_campaigns(df, latest_timestamp)

    summary = {
        "client_number": CLIENT_NUMBER,
        "report_suffix": REPORT_SUFFIX,
        "report_timezone": REPORT_TIMEZONE,
        "latest_timestamp": latest["timestamp"].strftime("%Y-%m-%d %H:%M:%S"),
        "previous_timestamp": previous["timestamp"].strftime("%Y-%m-%d %H:%M:%S"),
        "health_status": health_status,
        "source": "simulated_data",
        "requested_data_source": DATA_SOURCE,
        "mode": "technical_demo",
        "thresholds": {
            "cpa_threshold": CPA_THRESHOLD,
            "roas_threshold": ROAS_THRESHOLD,
            "ctr_threshold": CTR_THRESHOLD,
        },
        "latest": {
            "spend": latest["spend"],
            "impressions": latest["impressions"],
            "clicks": latest["clicks"],
            "ctr": latest["ctr"],
            "cpc": latest["cpc"],
            "conversions": latest["conversions"],
            "cvr": latest["cvr"],
            "cpa": latest["cpa"],
            "revenue": latest["revenue"],
            "roas": latest["roas"],
        },
        "previous": {
            "spend": previous["spend"],
            "impressions": previous["impressions"],
            "clicks": previous["clicks"],
            "ctr": previous["ctr"],
            "cpc": previous["cpc"],
            "conversions": previous["conversions"],
            "cvr": previous["cvr"],
            "cpa": previous["cpa"],
            "revenue": previous["revenue"],
            "roas": previous["roas"],
        },
        "changes": {
            "spend_change_pct": spend_change,
            "clicks_change_pct": clicks_change,
            "conversions_change_abs": conversions_change,
            "cpa_change_pct": cpa_change,
            "roas_change_pct": roas_change,
        },
        "alerts": alerts,
        "connection_status": connection_status,
    }

    report = f"""# HMA — Reporte Horario de Campañas

**Cliente:** {CLIENT_NUMBER}  
**Fecha/Hora:** {latest["timestamp"].strftime("%Y-%m-%d %H:%M")}  
**Zona horaria:** {REPORT_TIMEZONE}  
**Estado general:** {health_status}  
**Fuente efectiva:** datos simulados  
**Fuente solicitada:** {DATA_SOURCE}  
**Modo:** demo técnica  

---

## 1. Resumen ejecutivo

El sistema ejecutó una revisión horaria de campañas publicitarias, calculó KPIs principales y generó alertas según umbrales configurados.

Esta versión todavía no usa datos reales de Google Ads o Meta Ads. El módulo actual simula datos con estructura similar a plataformas publicitarias. En producción, el simulador se reemplaza por conectores API.

---

## 2. KPIs principales — última hora

| Métrica | Valor |
|---|---:|
| Gasto | ${latest["spend"]:.2f} |
| Impresiones | {latest["impressions"]:,} |
| Clics | {latest["clicks"]:,} |
| CTR | {latest["ctr"]:.2f}% |
| CPC | ${latest["cpc"]:.2f} |
| Conversiones | {latest["conversions"]} |
| CVR | {latest["cvr"]:.2f}% |
| CPA | ${latest["cpa"]:.2f} |
| Revenue | ${latest["revenue"]:.2f} |
| ROAS | {latest["roas"]:.2f} |

---

## 3. Comparación vs hora anterior

| Métrica | Hora actual | Hora anterior | Variación |
|---|---:|---:|---:|
| Gasto | ${latest["spend"]:.2f} | ${previous["spend"]:.2f} | {spend_change:+.1f}% |
| Clics | {latest["clicks"]:,} | {previous["clicks"]:,} | {clicks_change:+.1f}% |
| Conversiones | {latest["conversions"]} | {previous["conversions"]} | {conversions_change:+.0f} |
| CPA | ${latest["cpa"]:.2f} | ${previous["cpa"]:.2f} | {cpa_change:+.1f}% |
| ROAS | {latest["roas"]:.2f} | {previous["roas"]:.2f} | {roas_change:+.1f}% |

---

## 4. Alertas

"""

    for alert in alerts:
        report += f"- {alert}\n"

    report += """

---

## 5. Breakdown por campaña — última hora

| Plataforma | Campaña | Gasto | Impresiones | Clics | CTR | Conv. | CPA | ROAS |
|---|---|---:|---:|---:|---:|---:|---:|---:|
"""

    for _, row in latest_campaigns.iterrows():
        report += (
            f"| {row['platform']} "
            f"| {row['campaign_name']} "
            f"| ${row['spend']:.2f} "
            f"| {int(row['impressions']):,} "
            f"| {int(row['clicks']):,} "
            f"| {row['ctr']:.2f}% "
            f"| {int(row['conversions'])} "
            f"| ${row['cpa']:.2f} "
            f"| {row['roas']:.2f} |\n"
        )

    report += f"""

---

## 6. Estado de conexión

| Componente | Estado |
|---|---|
| DATA_SOURCE solicitado | {DATA_SOURCE} |
| Fuente efectiva actual | simulated_data |
| Ejecución real de APIs | {connection_status["real_api_execution_enabled"]} |
| Google Ads listo | {connection_status["google_ads_ready"]} |
| Meta Ads listo | {connection_status["meta_ads_ready"]} |

### Credenciales faltantes — Google Ads

"""

    if connection_status["missing_google_ads_credentials"]:
        for item in connection_status["missing_google_ads_credentials"]:
            report += f"- {item}\n"
    else:
        report += "- Ninguna. Credenciales mínimas detectadas.\n"

    report += "\n### Credenciales faltantes — Meta Ads\n\n"

    if connection_status["missing_meta_ads_credentials"]:
        for item in connection_status["missing_meta_ads_credentials"]:
            report += f"- {item}\n"
    else:
        report += "- Ninguna. Credenciales mínimas detectadas.\n"

    report += "\n### Notas de conexión\n\n"

    for note in connection_status["notes"]:
        report += f"- {note}\n"

    report += f"""

---

## 7. Interpretación operativa

- Si el CPA sube por encima de ${CPA_THRESHOLD:.2f}, el sistema marca alerta de eficiencia.
- Si el ROAS cae por debajo de {ROAS_THRESHOLD:.2f}, el sistema marca alerta de rentabilidad.
- Si hay gasto sin conversiones, el sistema recomienda revisar tracking, oferta, segmentación o calidad del tráfico.
- Si el CTR cae por debajo de {CTR_THRESHOLD:.2f}%, el sistema marca posible fatiga creativa o baja relevancia del anuncio.

---

## 8. Próxima fase

Para convertir esta demo en sistema real se necesita reemplazar `simulate_ads_data()` por conectores reales:

- Google Ads API
- Meta Marketing API

También se requiere definir:

- cuentas publicitarias;
- campañas a monitorear;
- zona horaria;
- umbrales por cliente;
- destino de alertas;
- almacenamiento persistente.

---

## 9. Estado técnico

| Componente | Estado |
|---|---|
| Ejecución local | OK |
| GitHub Actions | OK |
| Reporte Markdown | OK |
| Artifact descargable | OK |
| CSV exportable | OK |
| JSON exportable | OK |
| Estado de conexión API | OK |
| Datos reales API | Pendiente |
| Persistencia histórica externa | Pendiente |
| Alertas externas | Pendiente |

---

**Sufijo de reporte:** {REPORT_SUFFIX}

"""

    return report, summary, latest_campaigns


def write_exports(
    summary: dict,
    latest_campaigns: pd.DataFrame,
    connection_status: dict,
) -> tuple[Path, Path, Path]:
    """Genera CSV y JSON para uso estructurado."""
    latest_metrics_path = EXPORTS_DIR / "latest_metrics.csv"
    latest_summary_path = EXPORTS_DIR / "latest_summary.json"
    connection_status_path = EXPORTS_DIR / "connection_status.json"

    export_df = latest_campaigns.copy()
    export_df["timestamp"] = export_df["timestamp"].astype(str)

    export_df.to_csv(latest_metrics_path, index=False, encoding="utf-8")

    latest_summary_path.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    connection_status_path.write_text(
        json.dumps(connection_status, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    return latest_metrics_path, latest_summary_path, connection_status_path


def main() -> None:
    logging.info("Inicio de ejecución HMA demo profesional.")

    try:
        run_time = get_local_now()
        report_basename = build_report_basename(run_time)

        connection_status = build_connection_status()

        raw_df = simulate_ads_data()
        kpi_df = calculate_kpis(raw_df)

        report, summary, latest_campaigns = generate_markdown_report(
            df=kpi_df,
            connection_status=connection_status,
        )

        latest_report_path = REPORTS_DIR / "latest_hourly_report.md"
        named_report_path = REPORTS_DIR / f"{report_basename}.md"
        basename_marker_path = REPORTS_DIR / "report_basename.txt"

        latest_report_path.write_text(report, encoding="utf-8")
        named_report_path.write_text(report, encoding="utf-8")
        basename_marker_path.write_text(report_basename, encoding="utf-8")

        latest_metrics_path, latest_summary_path, connection_status_path = write_exports(
            summary=summary,
            latest_campaigns=latest_campaigns,
            connection_status=connection_status,
        )

        logging.info("Reporte generado correctamente.")
        logging.info(f"Reporte latest: {latest_report_path}")
        logging.info(f"Reporte nombrado: {named_report_path}")
        logging.info(f"CSV exportado: {latest_metrics_path}")
        logging.info(f"JSON summary exportado: {latest_summary_path}")
        logging.info(f"JSON connection status exportado: {connection_status_path}")
        logging.info(f"Nombre base de artifact: {report_basename}")
        logging.info(f"Zona horaria del reporte: {REPORT_TIMEZONE}")
        logging.info(f"Fuente solicitada: {DATA_SOURCE}")
        logging.info("Fuente efectiva: simulated_data")

        print(report)
        print(f"\nReporte latest generado en: {latest_report_path}")
        print(f"Reporte nombrado generado en: {named_report_path}")
        print(f"CSV generado en: {latest_metrics_path}")
        print(f"JSON summary generado en: {latest_summary_path}")
        print(f"JSON connection status generado en: {connection_status_path}")
        print(f"Nombre base de artifact: {report_basename}")
        print(f"Zona horaria del reporte: {REPORT_TIMEZONE}")
        print(f"Fuente solicitada: {DATA_SOURCE}")
        print("Fuente efectiva: simulated_data")

    except Exception as exc:
        logging.exception(f"Error durante ejecución HMA: {exc}")
        raise


if __name__ == "__main__":
    main()