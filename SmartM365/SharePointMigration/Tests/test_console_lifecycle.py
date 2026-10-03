"""Offline console lifecycle and Python CLI source checks."""

import ast
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
import os
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "Scripts"))
from console_lifecycle import run_console_script  # noqa: E402


class ConsoleLifecycleTests(unittest.TestCase):
    def test_success_and_stdout_contract(self):
        before = os.environ.pop("SPMIG_CONSOLE_LIFECYCLE_ACTIVE", None)
        stdout, stderr = StringIO(), StringIO()
        try:
            with redirect_stdout(stdout), redirect_stderr(stderr):
                run_console_script(lambda: print("DATA"), "synthetic.py", "1.0.0")
            self.assertEqual(stdout.getvalue(), "DATA\n")
            self.assertIn("SmartM365 by WorkplaceCloudHub", stderr.getvalue())
            self.assertIn("Status    : SUCCESS", stderr.getvalue())
            self.assertNotIn("SPMIG_CONSOLE_LIFECYCLE_ACTIVE", os.environ)
        finally:
            if before is not None:
                os.environ["SPMIG_CONSOLE_LIFECYCLE_ACTIVE"] = before

    def test_failure_and_nested_child(self):
        previous = os.environ.get("SPMIG_CONSOLE_LIFECYCLE_ACTIVE")
        stderr = StringIO()
        try:
            with redirect_stderr(stderr):
                with self.assertRaisesRegex(ValueError, "synthetic failure"):
                    run_console_script(lambda: (_ for _ in ()).throw(ValueError("synthetic failure")), "synthetic.py", "1.0.0")
            self.assertIn("Status    : FAILED", stderr.getvalue())
            os.environ["SPMIG_CONSOLE_LIFECYCLE_ACTIVE"] = "1"
            nested = StringIO()
            with redirect_stderr(nested):
                run_console_script(lambda: None, "child.py", "1.0.0")
            self.assertEqual(nested.getvalue(), "")
        finally:
            if previous is None:
                os.environ.pop("SPMIG_CONSOLE_LIFECYCLE_ACTIVE", None)
            else:
                os.environ["SPMIG_CONSOLE_LIFECYCLE_ACTIVE"] = previous

    def test_cli_files_parse(self):
        files = list((ROOT / "Scripts").rglob("*.py")) + [ROOT / "Tools" / "build_release.py"]
        for path in files:
            with self.subTest(path=path.name):
                ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
                if path.name not in ("console_lifecycle.py", "report_html.py"):
                    self.assertIn("run_console_script(", path.read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
