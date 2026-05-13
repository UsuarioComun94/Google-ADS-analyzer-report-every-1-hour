import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo


GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_REPOSITORY = os.getenv(
    "GITHUB_REPOSITORY",
    "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour",
)
GITHUB_RUN_ID = str(os.getenv("GITHUB_RUN_ID", "")).strip()
GITHUB_WORKFLOW_FILE = os.getenv("GITHUB_WORKFLOW_FILE", "hma-hourly.yml")
GITHUB_REF_NAME = os.getenv("GITHUB_REF_NAME", "main")
REPORT_TIMEZONE = os.getenv("REPORT_TIMEZONE", "America/Argentina/Cordoba")
CLIENT_NUMBER = os.getenv("CLIENT_NUMBER", "CLIENTE-DEMO-0001")
REPORT_SUFFIX = os.getenv("REPORT_SUFFIX", "JPPQ")


def set_output(name: str, value: str) -> None:
    output_path = os.getenv("GITHUB_OUTPUT")

    if output_path:
        with open(output_path, "a", encoding="utf-8") as output_file:
            output_file.write(f"{name}={value}\n")
    else:
        print(f"{name}={value}")


def safe_one_line(value: str) -> str:
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def target_context() -> dict:
    timezone = ZoneInfo(REPORT_TIMEZONE)
    now = datetime.now(timezone)
    hour_start = now.replace(minute=0, second=0, microsecond=0)
    hour_end = hour_start + timedelta(hours=1)

    date_slug = hour_start.strftime("%Y-%m-%d")
    hour_slug = hour_start.strftime("%H-00")
    target_hour = hour_start.isoformat(timespec="minutes")
    artifact_name = f"Campania_{CLIENT_NUMBER}_{date_slug}_{hour_slug}_{REPORT_SUFFIX}"

    return {
        "timezone": timezone,
        "hour_start": hour_start,
        "hour_end": hour_end,
        "target_hour": target_hour,
        "target_hour_slug": f"{date_slug}_{hour_slug}",
        "artifact_name": artifact_name,
    }


def finish(should_run: bool, reason: str, ctx: dict) -> None:
    clean_reason = safe_one_line(reason)

    set_output("should_run", "true" if should_run else "false")
    set_output("reason", clean_reason)
    set_output("target_hour", ctx["target_hour"])
    set_output("target_hour_slug", ctx["target_hour_slug"])
    set_output("artifact_name", ctx["artifact_name"])

    print(f"should_run={should_run}")
    print(f"target_hour={ctx['target_hour']}")
    print(f"target_hour_slug={ctx['target_hour_slug']}")
    print(f"artifact_name={ctx['artifact_name']}")
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


def list_recent_runs() -> list[dict]:
    encoded_workflow = urllib.parse.quote(GITHUB_WORKFLOW_FILE, safe="")
    encoded_branch = urllib.parse.quote(GITHUB_REF_NAME, safe="")

    path = (
        f"/repos/{GITHUB_REPOSITORY}/actions/workflows/"
        f"{encoded_workflow}/runs?per_page=30&branch={encoded_branch}"
    )

    data = github_api("GET", path)
    return data.get("workflow_runs", [])


def list_recent_artifacts() -> list[dict]:
    path = f"/repos/{GITHUB_REPOSITORY}/actions/artifacts?per_page=100"
    data = github_api("GET", path)
    return data.get("artifacts", [])


def is_hma_report_artifact(name: str) -> bool:
    if not name:
        return False

    if name.startswith("HMA_FAILURE_"):
        return False

    expected_prefix = f"Campania_{CLIENT_NUMBER}_"
    expected_suffix = f"_{REPORT_SUFFIX}"

    return name.startswith(expected_prefix) and name.endswith(expected_suffix)


def has_existing_artifact_for_target_hour(ctx: dict) -> tuple[bool, str]:
    artifacts = list_recent_artifacts()

    for artifact in artifacts:
        if artifact.get("expired") is True:
            continue

        name = str(artifact.get("name", "")).strip()
        if not is_hma_report_artifact(name):
            continue

        if name == ctx["artifact_name"]:
            return True, f"Ya existe artifact determinístico para la hora objetivo: {name}"

        created_at_raw = artifact.get("created_at")
        if not created_at_raw:
            continue

        created_at_local = parse_github_datetime(created_at_raw).astimezone(ctx["timezone"])

        if ctx["hour_start"] <= created_at_local < ctx["hour_end"]:
            run_id = ""
            workflow_run = artifact.get("workflow_run") or {}
            if isinstance(workflow_run, dict):
                run_id = str(workflow_run.get("id", ""))

            return (
                True,
                (
                    "Ya existe artifact HMA creado dentro de la hora objetivo. "
                    f"name={name}, run_id={run_id}, created_at={created_at_local.isoformat()}"
                ),
            )

    return False, ""


def has_active_run_for_target_hour(ctx: dict) -> tuple[bool, str]:
    runs = list_recent_runs()

    for run in runs:
        run_id = str(run.get("id", "")).strip()

        if not run_id or run_id == GITHUB_RUN_ID:
            continue

        created_at_raw = run.get("created_at")
        if not created_at_raw:
            continue

        created_at_local = parse_github_datetime(created_at_raw).astimezone(ctx["timezone"])

        if not (ctx["hour_start"] <= created_at_local < ctx["hour_end"]):
            continue

        status = run.get("status")
        event = run.get("event")
        conclusion = run.get("conclusion")

        if status in {"queued", "in_progress"}:
            return (
                True,
                (
                    "Ya existe un run activo para esta hora objetivo. "
                    f"id={run_id}, event={event}, status={status}, "
                    f"created_at={created_at_local.isoformat()}"
                ),
            )

        print(
            "Run no bloqueante dentro de la hora objetivo. "
            f"id={run_id}, event={event}, status={status}, conclusion={conclusion}, "
            f"created_at={created_at_local.isoformat()}"
        )

    return False, ""


def main() -> None:
    ctx = target_context()

    print(f"Target hour: {ctx['target_hour']}")
    print(f"Target hour slug: {ctx['target_hour_slug']}")
    print(f"Expected artifact name: {ctx['artifact_name']}")
    print(f"Current run id: {GITHUB_RUN_ID}")
    print(f"Workflow file: {GITHUB_WORKFLOW_FILE}")
    print(f"Repository: {GITHUB_REPOSITORY}")

    try:
        artifact_exists, artifact_reason = has_existing_artifact_for_target_hour(ctx)

        if artifact_exists:
            finish(False, artifact_reason, ctx)
            return

        active_run_exists, active_run_reason = has_active_run_for_target_hour(ctx)

        if active_run_exists:
            finish(False, active_run_reason, ctx)
            return

        finish(True, "No existe artifact HMA válido para esta hora objetivo. Se permite generar.", ctx)

    except Exception as exc:
        print(f"ERROR en gate: {exc}", file=sys.stderr)
        finish(
            False,
            (
                "Gate falló y se bloquea generación para evitar duplicados. "
                f"Error: {exc}"
            ),
            ctx,
        )


if __name__ == "__main__":
    main()
