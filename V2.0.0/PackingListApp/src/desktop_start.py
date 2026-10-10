"""Console-free packaged entry point with a readable startup error log."""
from pathlib import Path
from contextlib import redirect_stderr, redirect_stdout
import sys
import traceback


def start():
    root = Path(__file__).resolve().parents[2]
    log_path = root / 'Startup.log'
    try:
        with log_path.open('a', encoding='utf-8', buffering=1) as log, redirect_stdout(log), redirect_stderr(log):
            try:
                from app import main
                main()
            except Exception:
                traceback.print_exc()
                raise
    except Exception as exc:
        if '--self-test' not in sys.argv:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None,
                f'Packing List Manager could not start.\n\n{exc}\n\nDetails: {log_path}',
                'Packing List Manager', 0x10)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(start())
