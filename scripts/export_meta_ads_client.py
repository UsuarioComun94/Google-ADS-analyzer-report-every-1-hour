from pathlib import Path
from datetime import datetime
import argparse
import csv
import os
import json
import requests

from dotenv import load_dotenv

def extract_action_value(actions, action_type):
    if not actions:
        return ""
    for item in actions:
        if item.get("action_type") == action_type:
            return item.get("value", "")
    return ""

def fetch_all_pages(url, params):
    data = []

    while url:
        r = requests.get(url, params=params, timeout=60)
        print("status_code:", r.status_code)
        r.raise_for_status()
        payload = r.json()
        data.extend(payload.get("data", []))

        next_url = payload.get("paging", {}).get("next")
        url = next_url
        params = None

    return data

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--client-dir", required=True)
    parser.add_argument("--date-preset", default="today")
    args = parser.parse_args()

    client_dir = Path(args.client_dir)
    secrets_dir = client_dir / "secrets"
    export_dir = client_dir / "raw_exports" / "meta_ads"
    export_dir.mkdir(parents=True, exist_ok=True)

    env_path = secrets_dir / "meta_ads.env"

    if not env_path.exists():
        raise RuntimeError(f"No existe {env_path}")

    load_dotenv(env_path)

    token = os.getenv("META_ACCESS_TOKEN")
    ad_account_id = os.getenv("META_AD_ACCOUNT_ID")
    api_version = os.getenv("META_API_VERSION", "v21.0")

    if not token:
        raise RuntimeError("Falta META_ACCESS_TOKEN")
    if not ad_account_id:
        raise RuntimeError("Falta META_AD_ACCOUNT_ID")

    url = f"https://graph.facebook.com/{api_version}/{ad_account_id}/insights"
    params = {
        "level": "campaign",
        "fields": "date_start,date_stop,campaign_id,campaign_name,spend,impressions,clicks,actions,action_values",
        "date_preset": args.date_preset,
        "time_increment": "1",
        "access_token": token,
    }

    rows = fetch_all_pages(url, params)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out = export_dir / f"meta_ads_campaign_daily_{stamp}.csv"

    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "extracted_at",
            "source",
            "date_start",
            "date_stop",
            "campaign_id",
            "campaign_name",
            "spend",
            "impressions",
            "clicks",
            "conversions",
            "conversion_value",
            "raw_actions",
            "raw_action_values",
        ])

        for row in rows:
            actions = row.get("actions", [])
            action_values = row.get("action_values", [])

            conversions = (
                extract_action_value(actions, "purchase")
                or extract_action_value(actions, "lead")
                or extract_action_value(actions, "offsite_conversion.fb_pixel_purchase")
                or extract_action_value(actions, "offsite_conversion.fb_pixel_lead")
            )

            conversion_value = (
                extract_action_value(action_values, "purchase")
                or extract_action_value(action_values, "offsite_conversion.fb_pixel_purchase")
            )

            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                "meta_ads",
                row.get("date_start", ""),
                row.get("date_stop", ""),
                row.get("campaign_id", ""),
                row.get("campaign_name", ""),
                row.get("spend", ""),
                row.get("impressions", ""),
                row.get("clicks", ""),
                conversions,
                conversion_value,
                json.dumps(actions, ensure_ascii=False),
                json.dumps(action_values, ensure_ascii=False),
            ])

    print("META_EXPORT_OK")
    print(f"file={out}")
    print(f"rows={len(rows)}")

if __name__ == "__main__":
    main()
