import os, sys
import traceback


def get_log_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "error.log")

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