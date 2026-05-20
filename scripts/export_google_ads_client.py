from pathlib import Path
from datetime import datetime
import argparse
import csv
import os

from dotenv import load_dotenv
from google.ads.googleads.client import GoogleAdsClient

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--client-dir", required=True)
    parser.add_argument("--date-range", default="TODAY")
    args = parser.parse_args()

    client_dir = Path(args.client_dir)
    secrets_dir = client_dir / "secrets"
    export_dir = client_dir / "raw_exports" / "google_ads"
    export_dir.mkdir(parents=True, exist_ok=True)

    config_path = secrets_dir / "google-ads.yaml"
    env_path = secrets_dir / "google_ads_account.env"

    if not config_path.exists():
        raise RuntimeError(f"No existe {config_path}")
    if not env_path.exists():
        raise RuntimeError(f"No existe {env_path}")

    load_dotenv(env_path)
    customer_id = os.getenv("GOOGLE_ADS_CUSTOMER_ID")

    if not customer_id:
        raise RuntimeError("Falta GOOGLE_ADS_CUSTOMER_ID")

    client = GoogleAdsClient.load_from_storage(str(config_path))
    ga_service = client.get_service("GoogleAdsService")

    query = f"""
        SELECT
          segments.date,
          segments.hour,
          campaign.id,
          campaign.name,
          metrics.cost_micros,
          metrics.impressions,
          metrics.clicks,
          metrics.conversions,
          metrics.conversions_value
        FROM campaign
        WHERE segments.date DURING {args.date_range}
        ORDER BY segments.date, segments.hour, campaign.id
    """

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out = export_dir / f"google_ads_campaign_hourly_{stamp}.csv"

    rows_count = 0

    with out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "extracted_at",
            "source",
            "date",
            "hour",
            "campaign_id",
            "campaign_name",
            "spend",
            "impressions",
            "clicks",
            "conversions",
            "conversion_value",
        ])

        response = ga_service.search(customer_id=customer_id, query=query)

        for row in response:
            writer.writerow([
                datetime.now().isoformat(timespec="seconds"),
                "google_ads",
                row.segments.date,
                row.segments.hour,
                row.campaign.id,
                row.campaign.name,
                row.metrics.cost_micros / 1_000_000,
                row.metrics.impressions,
                row.metrics.clicks,
                row.metrics.conversions,
                row.metrics.conversions_value,
            ])
            rows_count += 1

    print("GOOGLE_EXPORT_OK")
    print(f"file={out}")
    print(f"rows={rows_count}")

if __name__ == "__main__":
    main()
