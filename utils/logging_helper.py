import os, sys, traceback, threading
from datetime import datetime

from utils.file_utils import get_log_path


def log_exception(error: BaseException | str, operation: str = "") -> None:
    if isinstance(error, str):
        error = RuntimeError(error)
    log_exc_triple(type(error), error, error.__traceback__, operation=operation)


def log_exc_triple(exc_type, exc_value, exc_tb, operation: str = ""):
    log_path = get_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(_format_exception_entry(exc_type, exc_value, exc_tb, operation))
            f.write("\n")
    except Exception:
        pass


def log_line(msg: str):
    """Write a non-exception diagnostic line. Prefer log_exception for failures."""
    with open(get_log_path(), "a", encoding="utf-8") as f:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"{timestamp} {msg}\n")


def delete_old_error_log():
    log_path = get_log_path()
    if os.path.exists(log_path):
        try:
            os.remove(log_path)
        except Exception:
            open(log_path, "w").close()


def _format_exception_entry(
    exc_type, exc_value, exc_tb, operation: str = ""
) -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    thread_name = threading.current_thread().name
    type_name = getattr(exc_type, "__name__", str(exc_type))
    message = str(exc_value)
    op = operation or "-"
    lines = [
        f"=== {timestamp} [{thread_name}] {op} ===",
        f"Type: {type_name}",
        f"Message: {message}",
        "Traceback:",
    ]
    if exc_tb is not None:
        lines.append(
            "".join(traceback.format_exception(exc_type, exc_value, exc_tb)).rstrip()
        )
    else:
        lines.append(
            "".join(traceback.format_exception_only(exc_type, exc_value)).rstrip()
        )
    lines.append("")
    return "\n".join(lines)


def _thread_excepthook(args):
    log_exc_triple(
        args.exc_type,
        args.exc_value,
        args.exc_traceback,
        operation="thread",
    )


sys.excepthook = log_exc_triple
threading.excepthook = _thread_excepthook
