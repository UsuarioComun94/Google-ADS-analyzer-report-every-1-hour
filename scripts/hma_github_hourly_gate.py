import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from zoneinfo import ZoneInfo


GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPOSITORY = os.getenv("GITHUB_REPOSITORY", "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour")
GITHUB_RUN_ID = str(os.getenv("GITHUB_RUN_ID", "")).strip()
GITHUB_WORKFLOW_FILE = os.getenv("GITHUB_WORKFLOW_FILE", "hma-hourly.yml")
GITHUB_REF_NAME = os.getenv("GITHUB_REF_NAME", "main")
REPORT_TIMEZONE = os.getenv("REPORT_TIMEZONE", "America/Argentina/Cordoba")


def set_output(name: str, value: str) -> None:
    output_path = os.getenv("GITHUB_OUTPUT")

    if output_path:
        with open(output_path, "a", encoding="utf-8") as output_file:
            output_file.write(f"{name}={value}\n")
    else:
        print(f"{name}={value}")


def finish(should_run: bool, reason: str, target_hour: str) -> None:
    clean_reason = reason.replace("\n", " ").replace("\r", " ")
    set_output("should_run", "true" if should_run else "false")
    set_output("reason", clean_reason)
    set_output("target_hour", target_hour)
    print(f"should_run={should_run}")
    print(f"target_hour={target_hour}")
    print(f"reason={clean_reason}")


def github_api(method: str, path: str) -> dict:
    if not GITHUB_TOKEN:
        raise RuntimeError("Falta GITHUB_TOKEN.")

    url = f"https://api.github.com{path}"

    request = urllib.request.Request(
        url=url,
        method=method,
        headers={
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "hma-github-hourly-gate",
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            raw_body = response.read().decode("utf-8", errors="replace")
            if not raw_body:
                return {}
            return json.loads(raw_body)

    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"GitHub API HTTP {exc.code}: {body}") from exc


def parse_github_datetime(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def current_hour_start() -> datetime:
    timezone = ZoneInfo(REPORT_TIMEZONE)
    now = datetime.now(timezone)
    return now.replace(minute=0, second=0, microsecond=0)


def artifact_count_for_run(run_id: str) -> int:
    path = f"/repos/{GITHUB_REPOSITORY}/actions/runs/{run_id}/artifacts"
    data = github_api("GET", path)
    return int(data.get("total_count", 0))


def list_recent_runs() -> list[dict]:
    encoded_workflow = urllib.parse.quote(GITHUB_WORKFLOW_FILE, safe="")
    encoded_branch = urllib.parse.quote(GITHUB_REF_NAME, safe="")

    path = (
        f"/repos/{GITHUB_REPOSITORY}/actions/workflows/"
        f"{encoded_workflow}/runs?per_page=30&branch={encoded_branch}"
    )

    data = github_api("GET", path)
    return data.get("workflow_runs", [])


def main() -> None:
    try:
        hour_start = current_hour_start()
        timezone = hour_start.tzinfo
        target_hour = hour_start.isoformat(timespec="minutes")

        print(f"Target hour: {target_hour}")
        print(f"Run actual: {GITHUB_RUN_ID}")
        print(f"Workflow file: {GITHUB_WORKFLOW_FILE}")

        runs = list_recent_runs()

        for run in runs:
            run_id = str(run.get("id", "")).strip()

            if not run_id or run_id == GITHUB_RUN_ID:
                continue

            created_at_raw = run.get("created_at")
            if not created_at_raw:
                continue

            created_at_local = parse_github_datetime(created_at_raw).astimezone(timezone)

            if created_at_local < hour_start:
                continue

            status = run.get("status")
            conclusion = run.get("conclusion")
            event = run.get("event")

            if status in {"queued", "in_progress"}:
                finish(
                    False,
                    (
                        "A run already exists in queued/in_progress state for this target hour. "
                        f"id={run_id}, event={event}, created_at={created_at_local.isoformat()}"
                    ),
                    target_hour,
                )
                return

            if status == "completed" and conclusion == "success":
                artifacts = artifact_count_for_run(run_id)

                if artifacts > 0:
                    finish(
                        False,
                        (
                            "A successful run with artifact already exists for this target hour. "
                            f"id={run_id}, event={event}, artifacts={artifacts}, "
                            f"created_at={created_at_local.isoformat()}"
                        ),
                        target_hour,
                    )
                    return

                print(
                    "Successful run without artifact inside target hour; retry is allowed. "
                    f"id={run_id}, event={event}, created_at={created_at_local.isoformat()}"
                )

            else:
                print(
                    "Non-blocking run inside target hour. "
                    f"id={run_id}, status={status}, conclusion={conclusion}, "
                    f"event={event}, created_at={created_at_local.isoformat()}"
                )

        finish(True, "No valid artifact exists for this target hour. Report generation is allowed.", target_hour)

    except Exception as exc:
        hour_start = current_hour_start()
        target_hour = hour_start.isoformat(timespec="minutes")
        print(f"ERROR en gate: {exc}", file=sys.stderr)
        finish(True, f"Gate failed; generation allowed to avoid missing target hour. Error: {exc}", target_hour)


if __name__ == "__main__":
    main()
