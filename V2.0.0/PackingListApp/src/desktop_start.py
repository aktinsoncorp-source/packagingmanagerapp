"""Console-free packaged entry point with a readable startup error log."""
from pathlib import Path
from contextlib import redirect_stderr, redirect_stdout
import sys
import traceback


def start():
    from app_paths import default_data_dir
    log_path = Path.home() / 'Packing List Manager Startup.log'
    try:
        root = default_data_dir()
        root.mkdir(parents=True, exist_ok=True)
        log_path = root / 'Startup.log'
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
