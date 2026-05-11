from pathlib import Path
import pandas as pd
import datetime
import random


BASE_DIR = Path(__file__).resolve().parent.parent
REPORTS_DIR = BASE_DIR / "reports"
REPORTS_DIR.mkdir(exist_ok=True)


def simulate_ads_data():
    """Simula la extracción de datos de una API de Ads."""
    now = datetime.datetime.now()
    data = {
        "timestamp": [now - datetime.timedelta(hours=i) for i in range(5)],
        "spend": [random.uniform(10, 50) for _ in range(5)],
        "clicks": [random.randint(100, 500) for _ in range(5)],
        "conversions": [random.randint(1, 10) for _ in range(5)],
    }
    return pd.DataFrame(data)


def generate_hourly_report(df):
    """Procesa los datos y genera un reporte en Markdown."""
    latest = df.iloc[0]
    previous = df.iloc[1]

    cpa_latest = (
        latest["spend"] / latest["conversions"]
        if latest["conversions"] > 0
        else 0
    )

    cpa_previous = (
        previous["spend"] / previous["conversions"]
        if previous["conversions"] > 0
        else 0
    )

    change_cpa = (
        ((cpa_latest - cpa_previous) / cpa_previous) * 100
        if cpa_previous > 0
        else 0
    )

    spend_change = (
        ((latest["spend"] - previous["spend"]) / previous["spend"]) * 100
        if previous["spend"] > 0
        else 0
    )

    clicks_change = (
        ((latest["clicks"] - previous["clicks"]) / previous["clicks"]) * 100
        if previous["clicks"] > 0
        else 0
    )

    report = f"""# Reporte Horario de Campaña

**Fecha/Hora:** {latest["timestamp"].strftime("%Y-%m-%d %H:%M")}

## KPIs Principales

| Métrica | Valor Actual | vs. Hora Anterior |
| :--- | :--- | :--- |
| **Gasto** | ${latest["spend"]:.2f} | {spend_change:+.1f}% |
| **Clics** | {latest["clicks"]} | {clicks_change:+.1f}% |
| **Conversiones** | {latest["conversions"]} | {latest["conversions"] - previous["conversions"]:+d} |
| **CPA** | ${cpa_latest:.2f} | {change_cpa:+.1f}% |

## Alertas de Sistema

"""

    if cpa_latest > 25:
        report += (
            f"⚠️ **ALERTA:** El CPA actual (${cpa_latest:.2f}) "
            "supera el umbral de seguridad de $25.00.\n"
        )
    else:
        report += "✅ Rendimiento dentro de los parámetros normales.\n"

    return report


if __name__ == "__main__":
    df_ads = simulate_ads_data()
    final_report = generate_hourly_report(df_ads)

    report_path = REPORTS_DIR / "latest_hourly_report.md"

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(final_report)

    print(f"Reporte horario generado exitosamente en: {report_path}")