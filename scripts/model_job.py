"""Run a model helper directly (no shell child) with captured diagnostics.

The UI owns this process handle and enforces its end-to-end deadline. Individual
SDK requests are bounded separately; neither layer automatically retries.
"""
from __future__ import annotations

import argparse
from contextlib import redirect_stdout, redirect_stderr
import os
from pathlib import Path
import runpy
import sys
import traceback


def model_timeout_seconds() -> float:
    try:
        return max(15.0, min(90.0, float(os.environ.get("JRPG_MODEL_TIMEOUT_SECONDS", "75"))))
    except ValueError:
        return 75.0


def gemini_http_options(types):
    return types.HttpOptions(
        timeout=int(model_timeout_seconds() * 1000),
        retry_options=types.HttpRetryOptions(attempts=1),
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--stdout", required=True)
    parser.add_argument("--stderr", required=True)
    parser.add_argument("script")
    parser.add_argument("args", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    script = str(Path(args.script).resolve())
    sys.argv = [script, *args.args]
    sys.path.insert(0, str(Path(script).parent))
    # Keep these streams alive even if the target rewraps sys.stdout.buffer.
    with open(args.stdout, "w", encoding="utf-8") as out, open(args.stderr, "w", encoding="utf-8") as err:
        with redirect_stdout(out), redirect_stderr(err):
            try:
                runpy.run_path(script, run_name="__main__")
            except SystemExit as exc:
                return exc.code if isinstance(exc.code, int) else (0 if exc.code is None else 1)
            except BaseException:
                traceback.print_exc()
                return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
