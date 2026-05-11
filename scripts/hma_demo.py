from pathlib import Path
import datetime as dt
import random
import os
import logging

import pandas as pd
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "reports"
LOGS_DIR = BASE_DIR / "logs"

REPORTS_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

load_dotenv(BASE_DIR / ".env")

CPA_THRESHOLD = float(os.getenv("CPA_THRESHOLD", 25))
ROAS_THRESHOLD = float(os.getenv("ROAS_THRESHOLD", 1.5))
CTR_THRESHOLD = float(os.getenv("CTR_THRESHOLD", 1.0))

logging.basicConfig(
    filename=LOGS_DIR / "hma.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


def safe_div(numerator: float, denominator: float) -> float:
    """Evita división por cero."""
    return numerator / denominator if denominator else 0


def simulate_ads_data() -> pd.DataFrame:
    """
    Simula datos horarios de campañas.

    Esta función representa el módulo extractor.
    Más adelante se reemplaza por:
    - fetch_google_ads_data()
    - fetch_meta_ads_data()
    """
    now = dt.datetime.now().replace(minute=0, second=0, microsecond=0)

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
        "spend": hour_df["spend"].sum(),
        "impressions": hour_df["impressions"].sum(),
        "clicks": hour_df["clicks"].sum(),
        "conversions": hour_df["conversions"].sum(),
        "revenue": hour_df["revenue"].sum(),
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


def classify_health(latest: dict, alerts: list[str]) -> str:
    """Clasifica estado general del sistema."""
    critical_alerts = [alert for alert in alerts if "⚠️" in alert]

    if len(critical_alerts) >= 3:
        return "CRÍTICO"
    if len(critical_alerts) >= 1:
        return "ATENCIÓN"
    return "NORMAL"


def generate_markdown_report(df: pd.DataFrame) -> str:
    """Genera reporte Markdown final."""
    timestamps = sorted(df["timestamp"].unique(), reverse=True)

    latest_timestamp = timestamps[0]
    previous_timestamp = timestamps[1]

    latest = summarize_hour(df, latest_timestamp)
    previous = summarize_hour(df, previous_timestamp)

    alerts = generate_alerts(latest)
    health_status = classify_health(latest, alerts)

    spend_change = percentage_change(latest["spend"], previous["spend"])
    clicks_change = percentage_change(latest["clicks"], previous["clicks"])
    conversions_change = latest["conversions"] - previous["conversions"]
    cpa_change = percentage_change(latest["cpa"], previous["cpa"])
    roas_change = percentage_change(latest["roas"], previous["roas"])

    latest_campaigns = df[df["timestamp"] == latest_timestamp].copy()
    latest_campaigns = latest_campaigns.sort_values("spend", ascending=False)

    report = f"""# HMA — Reporte Horario de Campañas

**Fecha/Hora:** {latest["timestamp"].strftime("%Y-%m-%d %H:%M")}  
**Estado general:** {health_status}  
**Fuente actual:** datos simulados  
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

## 6. Interpretación operativa

- Si el CPA sube por encima de ${CPA_THRESHOLD:.2f}, el sistema marca alerta de eficiencia.
- Si el ROAS cae por debajo de {ROAS_THRESHOLD:.2f}, el sistema marca alerta de rentabilidad.
- Si hay gasto sin conversiones, el sistema recomienda revisar tracking, oferta, segmentación o calidad del tráfico.
- Si el CTR cae por debajo de {CTR_THRESHOLD:.2f}%, el sistema marca posible fatiga creativa o baja relevancia del anuncio.

---

## 7. Próxima fase

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

## 8. Estado técnico

| Componente | Estado |
|---|---|
| Ejecución local | OK |
| GitHub Actions | OK |
| Reporte Markdown | OK |
| Artifact descargable | OK |
| Datos reales API | Pendiente |
| Persistencia histórica externa | Pendiente |
| Alertas externas | Pendiente |

"""

    return report


def main() -> None:
    logging.info("Inicio de ejecución HMA demo profesional.")

    try:
        raw_df = simulate_ads_data()
        kpi_df = calculate_kpis(raw_df)

        report = generate_markdown_report(kpi_df)

        latest_report_path = REPORTS_DIR / "latest_hourly_report.md"
        timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        historical_report_path = REPORTS_DIR / f"hma_report_{timestamp}.md"

        latest_report_path.write_text(report, encoding="utf-8")
        historical_report_path.write_text(report, encoding="utf-8")

        logging.info("Reporte generado correctamente.")
        logging.info(f"Reporte latest: {latest_report_path}")
        logging.info(f"Reporte histórico: {historical_report_path}")

        print(report)
        print(f"\nReporte latest generado en: {latest_report_path}")
        print(f"Reporte histórico generado en: {historical_report_path}")

    except Exception as exc:
        logging.exception(f"Error durante ejecución HMA: {exc}")
        raise


if __name__ == "__main__":
    main()