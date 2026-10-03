"""Console lifecycle for directly executed SharePoint migration Python scripts."""

from datetime import datetime
import os
from pathlib import Path
import sys
from time import monotonic


_MARKER = "SPMIG_CONSOLE_LIFECYCLE_ACTIVE"
_LOG = "SPMIG_CONSOLE_LIFECYCLE_LOG"
_RULE = "=" * 80


def _write(line: str) -> None:
    print(line, file=sys.stderr, flush=True)


def run_console_script(main, script_path: str, version: str, action: str = ""):
    """Run a CLI with one branded start/end pair and unchanged stdout."""
    previous = os.environ.get(_MARKER)
    previous_log = os.environ.get(_LOG)
    owns_console = not previous
    started = datetime.now()
    started_clock = monotonic()
    name = Path(script_path).name
    action = action or Path(script_path).stem
    status = "SUCCESS"
    reason = ""
    if owns_console:
        os.environ[_MARKER] = "1"
        os.environ.pop(_LOG, None)
        stamp = started.strftime("%Y-%m-%d %H:%M:%S")
        _write(f"{stamp} Script  : {name} v{version}")
        for line in (
            _RULE,
            " SmartM365 by WorkplaceCloudHub",
            " Website : https://workplacecloudhub.com",
            " GitHub  : https://github.com/khda79/workplacecloudhub.com",
            _RULE,
        ):
            _write(line)
        _write(f"{stamp} Action    : {action}")
        _write(f"{stamp} Started   : {stamp}")
    try:
        result = main()
        if isinstance(result, int) and not isinstance(result, bool) and result != 0:
            status = "FAILED"
            reason = f"Exit code {result}"
        return result
    except KeyboardInterrupt:
        status = "CANCELLED"
        reason = "Interrupted by user"
        raise
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else (0 if exc.code is None else 1)
        if code != 0:
            status = "FAILED"
            reason = f"Exit code {code}"
        raise
    except BaseException as exc:
        status = "FAILED"
        reason = str(exc)
        raise
    finally:
        if owns_console:
            try:
                ended = datetime.now()
                stamp = ended.strftime("%Y-%m-%d %H:%M:%S")
                seconds = int(monotonic() - started_clock)
                duration = f"{seconds // 3600:02d}:{seconds // 60 % 60:02d}:{seconds % 60:02d}"
                lines = [
                    _RULE,
                    "SmartM365 by WorkplaceCloudHub - " + {
                        "SUCCESS": "Execution completed",
                        "FAILED": "Execution failed",
                        "CANCELLED": "Execution cancelled",
                    }[status],
                    f"Script    : {name}",
                    f"Action    : {action}",
                    f"Status    : {status}",
                    f"Duration  : {duration}",
                ]
                log_path = os.environ.get(_LOG, "")
                if log_path and Path(log_path).is_file():
                    lines.append(f"Log       : {log_path}")
                if reason:
                    lines.append(f"Reason    : {reason}")
                lines.append(_RULE)
                output = [f"{stamp} {line}" for line in lines]
                for line in output:
                    _write(line)
                if log_path and Path(log_path).is_file():
                    try:
                        with open(log_path, "a", encoding="utf-8") as log:
                            log.write("\n".join(output) + "\n")
                    except OSError as exc:
                        _write(f"{stamp} Could not append completion to log: {exc}")
            finally:
                if previous is None:
                    os.environ.pop(_MARKER, None)
                else:
                    os.environ[_MARKER] = previous
                if previous_log is None:
                    os.environ.pop(_LOG, None)
                else:
                    os.environ[_LOG] = previous_log
