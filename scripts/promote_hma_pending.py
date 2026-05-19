from pathlib import Path
from datetime import datetime
import sys
import ctypes
import atexit


BASE_DIR = Path(__file__).resolve().parent.parent
HISTORY_DIR = BASE_DIR / "historico"
MASTER_FILE = HISTORY_DIR / "HMA_Master.xlsx"
LOG_FILE = BASE_DIR / "logs" / "hma_pending_promotion.log"

_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_writer_mutex = _kernel32.CreateMutexW(None, False, "HMA_HourlyMarketingAnalyzer_Writer_Lock")
if ctypes.get_last_error() == 183:
    print("Otra ejecución escritora HMA ya está activa. Se omite promote_hma_pending.py.")
    sys.exit(0)

atexit.register(lambda: _kernel32.CloseHandle(_writer_mutex) if _writer_mutex else None)



def log(message: str) -> None:
    LOG_FILE.parent.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {message}"
    print(line)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def is_file_locked(path: Path) -> bool:
    if not path.exists():
        return False

    try:
        with path.open("a+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return True


def get_pending_files() -> list[Path]:
    if not HISTORY_DIR.exists():
        return []

    return sorted(
        HISTORY_DIR.glob("HMA_Master_PENDING_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


def promote_latest_pending() -> int:
    pending_files = get_pending_files()

    if not pending_files:
        log("No hay archivos PENDING para promover.")
        return 0

    latest_pending = pending_files[0]

    if is_file_locked(MASTER_FILE):
        log(f"HMA_Master.xlsx está abierto/bloqueado. No se puede promover todavía: {latest_pending.name}")
        return 2

    try:
        backup_file = HISTORY_DIR / f"HMA_Master_BACKUP_BEFORE_PENDING_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"

        if MASTER_FILE.exists():
            MASTER_FILE.replace(backup_file)
            log(f"Backup creado: {backup_file.name}")

        latest_pending.replace(MASTER_FILE)
        log(f"PENDING promovido a HMA_Master.xlsx: {latest_pending.name}")

        for old_pending in pending_files[1:]:
            try:
                old_pending.unlink()
                log(f"PENDING viejo eliminado: {old_pending.name}")
            except Exception as exc:
                log(f"No se pudo eliminar PENDING viejo {old_pending.name}: {exc}")

        return 0

    except PermissionError:
        log("PermissionError al promover. Excel probablemente sigue abierto.")
        return 2

    except Exception as exc:
        log(f"ERROR al promover PENDING: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(promote_latest_pending())
