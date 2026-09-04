# -*- coding: utf-8 -*-
"""Remove one JRPG Translator bridge workspace outside the UI process."""

from __future__ import annotations

import argparse
import re
import shutil
import time
from pathlib import Path


WORKSPACE_RE = re.compile(r"^[A-Za-z0-9_-]+_\d+_\d+_\d+$")


def validated_workspace(root_value: Path, target_value: Path) -> tuple[Path, Path]:
    root = root_value.resolve(strict=False)
    target = target_value.resolve(strict=False)
    if target.parent != root or not WORKSPACE_RE.fullmatch(target.name):
        raise ValueError("Refusing to remove a path outside the bridge workspace root")
    return root, target


def remove_workspace(root: Path, target: Path) -> bool:
    # Antivirus/indexing can transiently retain freshly written bridge files.
    # This helper may wait and retry because it is deliberately detached from
    # the AutoHotkey UI process.
    for delay in (0.0, 0.1, 0.25, 0.5, 1.0, 2.0):
        if delay:
            time.sleep(delay)
        try:
            shutil.rmtree(target)
        except FileNotFoundError:
            break
        except OSError:
            continue
        else:
            break
    if target.exists():
        return False
    try:
        root.rmdir()
    except OSError:
        pass
    return True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--target", required=True, type=Path)
    arguments = parser.parse_args()
    try:
        root, target = validated_workspace(arguments.root, arguments.target)
    except ValueError:
        return 2
    return 0 if remove_workspace(root, target) else 1


if __name__ == "__main__":
    raise SystemExit(main())
