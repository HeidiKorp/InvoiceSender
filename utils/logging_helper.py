import os, sys
import traceback

from utils.file_utils import get_log_path

def log_exc_triple(exc_type, exc_value, exc_tb):
    # Write errors to a log file next to the exe, even in PyInstaller
    log_path = get_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("=== Exception ===\n")
            traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
    except Exception:
        pass

sys.excepthook = log_exc_triple

def _thread_excepthook(args):
    log_exc_triple(args.exc_type, args.exc_value, args.exc_traceback)


def log_exception(e: Exception):
    exc_type = type(e)
    exc_value = e
    exc_tb = e.__traceback__
    log_exc_triple(exc_type, exc_value, exc_tb)


def delete_old_error_log():
    log_path = get_log_path()
    if os.path.exists(log_path):
        try:
            os.remove(log_path)
        except Exception:
            # If deletion fails, overwrite instead of crash
            open(log_path, "w").close()

def log_line(msg: str):
    with open(get_log_path(), "a", encoding="utf-8") as f:
        f.write(msg + "\n")