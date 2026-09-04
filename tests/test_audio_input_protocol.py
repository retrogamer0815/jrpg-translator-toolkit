"""Test diagnostic output without importing drivers, dotenv, or AI clients."""
import ast
import contextlib
import io
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch


SOURCE = Path(__file__).resolve().parents[1] / "scripts" / "live_audio_translator.py"
TREE = ast.parse(SOURCE.read_text(encoding="utf-8-sig"), filename=str(SOURCE))
NODES = [node for node in TREE.body if isinstance(node, ast.FunctionDef)
         and node.name in {"run_speaker_list", "emit_audio_test_result", "run_audio_test"}]
CODE = compile(ast.Module(body=NODES, type_ignores=[]), str(SOURCE), "exec")


class AudioProtocolTests(unittest.TestCase):
    def invoke(self, devices=None, error=None, output=True):
        def speakers():
            if error:
                raise RuntimeError(error)
            return [SimpleNamespace(name=name) for name in (devices or [])]

        namespace = {"os": os, "sys": sys, "sc": SimpleNamespace(all_speakers=speakers)}
        exec(CODE, namespace)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "devices.txt"
            with patch.dict(os.environ, {"AUDIO_DEVICE_LIST_RESULT_FILE": str(target) if output else ""}):
                stdout = io.StringIO()
                with contextlib.redirect_stdout(stdout):
                    result = namespace["run_speaker_list"]()
                return result, stdout.getvalue(), target.read_text(encoding="utf-8") if target.exists() else None

    def test_success_unicode_deduplication(self):
        status, stdout, result = self.invoke(["Speakers", "日本語 ＆ 音声", "Speakers", "  "])
        self.assertEqual(status, 0)
        self.assertEqual(stdout, "Speakers\n日本語 ＆ 音声\n")
        self.assertEqual(result, "JRPG_AUDIO_DEVICES:OK\nSpeakers\n日本語 ＆ 音声\n")

    def test_legacy_stdout_without_result_file(self):
        status, stdout, result = self.invoke(["Speakers", "Headphones"], output=False)
        self.assertEqual((status, stdout, result), (0, "Speakers\nHeadphones\n", None))

    def test_empty_list_is_explicit_success(self):
        status, stdout, result = self.invoke([])
        self.assertEqual(status, 0)
        self.assertEqual(stdout, "")
        self.assertEqual(result, "JRPG_AUDIO_DEVICES:OK\n\n")

    def test_driver_error_is_not_empty_success(self):
        status, stdout, result = self.invoke(error="Driver failed\r\nTry again")
        self.assertEqual(status, 1)
        self.assertEqual(stdout, "")
        self.assertEqual(result, "JRPG_AUDIO_DEVICES:ERROR\nDriver failed  Try again\n")

    def test_device_names_cannot_inject_protocol_lines(self):
        status, _, result = self.invoke(["Device\nJRPG_AUDIO_DEVICES:ERROR\rnext"])
        self.assertEqual(status, 0)
        self.assertEqual(result.splitlines(), ["JRPG_AUDIO_DEVICES:OK", "Device JRPG_AUDIO_DEVICES:ERROR next"])

    def test_output_write_failure(self):
        namespace = {"os": os, "sys": sys, "sc": SimpleNamespace(all_speakers=lambda: [])}
        exec(CODE, namespace)
        with tempfile.TemporaryDirectory() as directory:
            with patch.dict(os.environ, {"AUDIO_DEVICE_LIST_RESULT_FILE": directory}):
                self.assertEqual(namespace["run_speaker_list"](), 1)

    def test_existing_audio_test_protocol(self):
        for detected in [True, False]:
            namespace = {"os": os, "sys": sys,
                         "test_audio_input": lambda: (detected, 0.5, 0.1, "Synthetic output")}
            exec(CODE, namespace)
            with tempfile.TemporaryDirectory() as directory:
                target = Path(directory) / "audio.txt"
                with patch.dict(os.environ, {"AUDIO_TEST_RESULT_FILE": str(target)}):
                    with contextlib.redirect_stdout(io.StringIO()):
                        self.assertEqual(namespace["run_audio_test"](), 0 if detected else 2)
                self.assertTrue(target.read_text(encoding="utf-8").startswith(
                    "JRPG_AUDIO_TEST:DETECTED:" if detected else "JRPG_AUDIO_TEST:SILENT:"))


if __name__ == "__main__":
    unittest.main()
