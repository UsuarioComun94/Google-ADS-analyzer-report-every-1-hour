import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from zoneinfo import ZoneInfo


GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_OWNER = os.getenv("GITHUB_OWNER", "UsuarioComun94")
GITHUB_REPO = os.getenv("GITHUB_REPO", "Google-ADS-analyzer-report-every-1-hour")
GITHUB_WORKFLOW_FILE = os.getenv("GITHUB_WORKFLOW_FILE", "hma-hourly.yml")
GITHUB_REF = os.getenv("GITHUB_REF", "main")

REPORT_TIMEZONE = os.getenv("REPORT_TIMEZONE", "America/Argentina/Cordoba")


def fail(message: str, code: int = 1) -> None:
    print(f"ERROR: {message}")
    sys.exit(code)


def env_input(name: str, default: str) -> str:
    value = os.getenv(name, default)
    return str(value).strip()


def github_request(method: str, url: str, payload: dict | None = None) -> tuple[int, str]:
    if not GITHUB_TOKEN:
        fail("Falta GITHUB_TOKEN en variables de entorno.")

    data = None
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")

    request = urllib.request.Request(
        url=url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {GITHUB_TOKEN}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
            "User-Agent": "hma-render-cron",
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8", errors="replace")
            return response.status, body
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        fail(f"GitHub HTTP {exc.code}: {body}")
    except Exception as exc:
        fail(f"No se pudo llamar a GitHub API: {exc}")

    return 0, ""


def parse_github_datetime(value: str) -> datetime:
    # GitHub returns ISO timestamps ending in Z.
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def current_hour_start() -> datetime:
    try:
        tz = ZoneInfo(REPORT_TIMEZONE)
    except Exception:
        fail(f"Zona horaria inválida en REPORT_TIMEZONE: {REPORT_TIMEZONE}")

    now = datetime.now(tz)
    return now.replace(minute=0, second=0, microsecond=0)


def list_recent_runs() -> list[dict]:
    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/actions/workflows/"
        f"{GITHUB_WORKFLOW_FILE}/runs?per_page=30"
    )

    status, body = github_request("GET", url)

    if status != 200:
        fail(f"No se pudieron listar runs. HTTP {status}")

    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        fail("GitHub devolvió JSON inválido al listar runs.")

    return data.get("workflow_runs", [])


def has_valid_run_for_current_hour(runs: list[dict]) -> bool:
    hour_start_local = current_hour_start()
    tz = hour_start_local.tzinfo

    print(f"Hora actual evaluada desde: {hour_start_local.isoformat()}")

    for run in runs:
        created_at_raw = run.get("created_at")
        if not created_at_raw:
            continue

        created_at_local = parse_github_datetime(created_at_raw).astimezone(tz)
        status = run.get("status")
        conclusion = run.get("conclusion")
        event = run.get("event")
        run_id = run.get("id")

        if created_at_local < hour_start_local:
            continue

        # Si ya existe un run creado dentro de esta hora, no disparar otro.
        # Cuenta queued/in_progress/completed success como bloqueo válido.
        if status in {"queued", "in_progress"}:
            print(
                "SKIP: ya existe un run en cola/en curso para esta hora "
                f"(id={run_id}, event={event}, created_at={created_at_local.isoformat()})."
            )
            return True

        if status == "completed" and conclusion == "success":
            print(
                "SKIP: ya existe un run exitoso para esta hora "
                f"(id={run_id}, event={event}, created_at={created_at_local.isoformat()})."
            )
            return True

        # Si hay un run de esta hora fallido/cancelado, no bloquea.
        # Render puede volver a disparar para recuperar esa hora.
        print(
            "INFO: run de esta hora no bloqueante "
            f"(id={run_id}, status={status}, conclusion={conclusion}, event={event})."
        )

    return False


def dispatch_workflow() -> None:
    url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/actions/workflows/"
        f"{GITHUB_WORKFLOW_FILE}/dispatches"
    )

    payload = {
        "ref": GITHUB_REF,
        "inputs": {
            "client_number": env_input("CLIENT_NUMBER", "CLIENTE-DEMO-0001"),
            "report_suffix": env_input("REPORT_SUFFIX", "JPPQ"),
            "report_timezone": env_input("REPORT_TIMEZONE", REPORT_TIMEZONE),
            "data_source": env_input("DATA_SOURCE", "simulated"),
            "cpa_threshold": env_input("CPA_THRESHOLD", "25"),
            "roas_threshold": env_input("ROAS_THRESHOLD", "1.5"),
            "ctr_threshold": env_input("CTR_THRESHOLD", "1.0"),
            "daily_budget_usd": env_input("DAILY_BUDGET_USD", "50000"),
        },
    }

    print("Disparando workflow_dispatch...")
    print(f"Repo: {GITHUB_OWNER}/{GITHUB_REPO}")
    print(f"Workflow: {GITHUB_WORKFLOW_FILE}")
    print(f"Ref: {GITHUB_REF}")
    print(f"Inputs: {json.dumps(payload['inputs'], ensure_ascii=False)}")

    status, _body = github_request("POST", url, payload)

    if status == 204:
        print("OK: workflow_dispatch enviado correctamente.")
        return

    fail(f"Respuesta inesperada de GitHub: HTTP {status}")


def main() -> None:
    if not GITHUB_TOKEN:
        fail("Falta GITHUB_TOKEN en variables de entorno.")

    runs = list_recent_runs()

    if has_valid_run_for_current_hour(runs):
        print("No se dispara workflow nuevo para evitar duplicado horario.")
        return

    dispatch_workflow()


if __name__ == "__main__":
    main()
