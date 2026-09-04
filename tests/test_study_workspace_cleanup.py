from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from scripts.study_workspace_cleanup import remove_workspace, validated_workspace


class StudyWorkspaceCleanupTests(unittest.TestCase):
    def test_removes_only_valid_direct_child_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "JRPG_Translator_Study_Bridge"
            target = root / "anki_candidates_123_456_1"
            nested = target / "nested"
            nested.mkdir(parents=True)
            (nested / "result.tsv").write_text("test", encoding="utf-8")

            validated_root, validated_target = validated_workspace(root, target)
            self.assertTrue(remove_workspace(validated_root, validated_target))
            self.assertFalse(target.exists())

    def test_rejects_parent_and_unrelated_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "JRPG_Translator_Study_Bridge"
            root.mkdir()
            for unsafe in (
                root,
                Path(temporary),
                root / "ordinary-folder",
                root / ".." / "outside_123_456_1",
            ):
                with self.assertRaises(ValueError):
                    validated_workspace(root, unsafe)


if __name__ == "__main__":
    unittest.main()
