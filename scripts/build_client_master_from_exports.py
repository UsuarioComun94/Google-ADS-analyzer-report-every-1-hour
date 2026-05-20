from pathlib import Path
from datetime import datetime
import argparse
import json
import math

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule


def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        if isinstance(value, float) and math.isnan(value):
            return default
        if str(value).strip() == "":
            return default
        return float(value)
    except Exception:
        return default


def safe_div(a, b):
    a = safe_float(a)
    b = safe_float(b)
    if b == 0:
        return 0.0
    return a / b


def read_csv_folder(folder: Path) -> pd.DataFrame:
    files = sorted(folder.glob("*.csv"))
    frames = []

    for f in files:
        try:
            df = pd.read_csv(f)
            df["source_file"] = f.name
            frames.append(df)
        except Exception as exc:
            print(f"ERROR_READING_CSV file={f} error={exc}")

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def normalize_google(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["extracted_at"] = df.get("extracted_at", "")
    out["platform"] = "google_ads"
    out["date"] = df.get("date", "")
    out["hour"] = df.get("hour", "")
    out["campaign_id"] = df.get("campaign_id", "")
    out["campaign_name"] = df.get("campaign_name", "")
    out["spend"] = df.get("spend", 0).apply(safe_float)
    out["impressions"] = df.get("impressions", 0).apply(safe_float)
    out["clicks"] = df.get("clicks", 0).apply(safe_float)
    out["conversions"] = df.get("conversions", 0).apply(safe_float)
    out["conversion_value"] = df.get("conversion_value", 0).apply(safe_float)
    out["source_file"] = df.get("source_file", "")

    return out


def normalize_meta(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["extracted_at"] = df.get("extracted_at", "")
    out["platform"] = "meta_ads"
    out["date"] = df.get("date_start", "")
    out["hour"] = ""
    out["campaign_id"] = df.get("campaign_id", "")
    out["campaign_name"] = df.get("campaign_name", "")
    out["spend"] = df.get("spend", 0).apply(safe_float)
    out["impressions"] = df.get("impressions", 0).apply(safe_float)
    out["clicks"] = df.get("clicks", 0).apply(safe_float)
    out["conversions"] = df.get("conversions", 0).apply(safe_float)
    out["conversion_value"] = df.get("conversion_value", 0).apply(safe_float)
    out["source_file"] = df.get("source_file", "")

    return out


def add_calculated_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    df["ctr"] = df.apply(lambda r: safe_div(r["clicks"], r["impressions"]), axis=1)
    df["cvr"] = df.apply(lambda r: safe_div(r["conversions"], r["clicks"]), axis=1)
    df["cpa"] = df.apply(lambda r: safe_div(r["spend"], r["conversions"]), axis=1)
    df["roas"] = df.apply(lambda r: safe_div(r["conversion_value"], r["spend"]), axis=1)

    return df


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "platform", "date", "spend", "impressions", "clicks",
            "conversions", "conversion_value", "ctr", "cvr", "cpa", "roas"
        ])

    grouped = (
        df.groupby(["platform", "date"], dropna=False)
        .agg({
            "spend": "sum",
            "impressions": "sum",
            "clicks": "sum",
            "conversions": "sum",
            "conversion_value": "sum",
        })
        .reset_index()
    )

    grouped["ctr"] = grouped.apply(lambda r: safe_div(r["clicks"], r["impressions"]), axis=1)
    grouped["cvr"] = grouped.apply(lambda r: safe_div(r["conversions"], r["clicks"]), axis=1)
    grouped["cpa"] = grouped.apply(lambda r: safe_div(r["spend"], r["conversions"]), axis=1)
    grouped["roas"] = grouped.apply(lambda r: safe_div(r["conversion_value"], r["spend"]), axis=1)

    return grouped


def build_recommendations(summary: pd.DataFrame) -> pd.DataFrame:
    rows = []

    if summary.empty:
        return pd.DataFrame(columns=[
            "platform", "date", "priority", "signal", "recommendation", "reason"
        ])

    for _, r in summary.iterrows():
        platform = r.get("platform", "")
        date = r.get("date", "")
        spend = safe_float(r.get("spend"))
        clicks = safe_float(r.get("clicks"))
        conversions = safe_float(r.get("conversions"))
        ctr = safe_float(r.get("ctr"))
        cvr = safe_float(r.get("cvr"))
        cpa = safe_float(r.get("cpa"))
        roas = safe_float(r.get("roas"))

        if spend > 0 and clicks == 0:
            rows.append({
                "platform": platform,
                "date": date,
                "priority": "HIGH",
                "signal": "Spend sin clicks",
                "recommendation": "Revisar tracking, delivery, segmentacion o estado de campañas.",
                "reason": f"Spend={spend}, clicks=0"
            })

        if clicks > 0 and conversions == 0:
            rows.append({
                "platform": platform,
                "date": date,
                "priority": "MEDIUM",
                "signal": "Clicks sin conversiones",
                "recommendation": "Revisar landing, oferta, formulario, pixel/conversion tracking.",
                "reason": f"Clicks={clicks}, conversions=0"
            })

        if ctr > 0.05 and cvr < 0.01 and clicks >= 50:
            rows.append({
                "platform": platform,
                "date": date,
                "priority": "MEDIUM",
                "signal": "CTR alto + CVR bajo",
                "recommendation": "No escalar presupuesto hasta validar landing, mensaje y friccion de conversion.",
                "reason": f"CTR={ctr:.2%}, CVR={cvr:.2%}"
            })

        if cpa > 0 and roas == 0 and conversions > 0:
            rows.append({
                "platform": platform,
                "date": date,
                "priority": "MEDIUM",
                "signal": "Conversiones sin valor",
                "recommendation": "Revisar value tracking, conversion_value o configuracion de evento.",
                "reason": f"CPA={cpa:.2f}, ROAS={roas:.2f}"
            })

    if not rows:
        rows.append({
            "platform": "all",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "priority": "INFO",
            "signal": "Sin alertas criticas",
            "recommendation": "Mantener monitoreo y validar consistencia de datos.",
            "reason": "No se detectaron reglas criticas con los datos disponibles."
        })

    return pd.DataFrame(rows)


def write_excel(path: Path, sheets: dict):
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_name = name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

    style_workbook(path)


def style_workbook(path: Path):
    wb = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center")

        for col in ws.columns:
            max_len = 10
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                header = ws.cell(row=1, column=col).value
                cell = ws.cell(row=row, column=col)

                if header in ["spend", "cpa", "conversion_value"]:
                    cell.number_format = '$#,##0.00'
                elif header in ["ctr", "cvr", "roas"]:
                    cell.number_format = '0.00%'
                elif header in ["impressions", "clicks", "conversions", "CSV_Google", "CSV_Meta"]:
                    cell.number_format = '#,##0'

        if ws.title == "recommendations" and ws.max_row > 1:
            priority_col = None
            for idx, cell in enumerate(ws[1], start=1):
                if cell.value == "priority":
                    priority_col = idx
                    break

            if priority_col:
                col_letter = get_column_letter(priority_col)
                ws.conditional_formatting.add(
                    f"{col_letter}2:{col_letter}{ws.max_row}",
                    CellIsRule(operator="equal", formula=['"HIGH"'], fill=PatternFill("solid", fgColor="FCA5A5"))
                )
                ws.conditional_formatting.add(
                    f"{col_letter}2:{col_letter}{ws.max_row}",
                    CellIsRule(operator="equal", formula=['"MEDIUM"'], fill=PatternFill("solid", fgColor="FDE68A"))
                )
                ws.conditional_formatting.add(
                    f"{col_letter}2:{col_letter}{ws.max_row}",
                    CellIsRule(operator="equal", formula=['"INFO"'], fill=PatternFill("solid", fgColor="BFDBFE"))
                )

    wb.save(path)


def load_client_config(client_dir: Path) -> dict:
    config_path = client_dir / "config" / "client_config.json"

    if not config_path.exists():
        return {}

    try:
        return json.loads(config_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--client-dir", required=True)
    args = parser.parse_args()

    client_dir = Path(args.client_dir)
    config = load_client_config(client_dir)

    client_id = config.get("client_id", client_dir.name)
    client_name = config.get("client_name", client_dir.name)

    google_raw = read_csv_folder(client_dir / "raw_exports" / "google_ads")
    meta_raw = read_csv_folder(client_dir / "raw_exports" / "meta_ads")

    google_norm = normalize_google(google_raw)
    meta_norm = normalize_meta(meta_raw)

    normalized = pd.concat([google_norm, meta_norm], ignore_index=True) if not google_norm.empty or not meta_norm.empty else pd.DataFrame()
    normalized = add_calculated_columns(normalized)

    summary = build_summary(normalized)
    recommendations = build_recommendations(summary)

    overview = pd.DataFrame([{
        "client_id": client_id,
        "client_name": client_name,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "google_csv_files": len(list((client_dir / "raw_exports" / "google_ads").glob("*.csv"))) if (client_dir / "raw_exports" / "google_ads").exists() else 0,
        "meta_csv_files": len(list((client_dir / "raw_exports" / "meta_ads").glob("*.csv"))) if (client_dir / "raw_exports" / "meta_ads").exists() else 0,
        "normalized_rows": len(normalized),
        "summary_rows": len(summary),
        "recommendation_rows": len(recommendations),
    }])

    output = client_dir / "historico" / "HMA_Master.xlsx"

    sheets = {
        "overview": overview,
        "normalized_metrics": normalized,
        "summary": summary,
        "recommendations": recommendations,
        "google_raw": google_raw,
        "meta_raw": meta_raw,
    }

    write_excel(output, sheets)

    print("CLIENT_MASTER_OK")
    print(f"client={client_id} | {client_name}")
    print(f"file={output}")
    print(f"normalized_rows={len(normalized)}")
    print(f"recommendations_rows={len(recommendations)}")


if __name__ == "__main__":
    main()
