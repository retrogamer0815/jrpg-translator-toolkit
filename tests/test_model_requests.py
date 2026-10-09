"""Offline request deadlines, diagnostic transport and provider-client contracts."""
import ast
import io
import os
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import example_sentence
import model_job


class ModelRequests(unittest.TestCase):
    def test_timeout_is_bounded_and_invalid_configuration_falls_back(self):
        for value, expected in [("75", 75), ("invalid", 75), ("900", 90), ("0", 15)]:
            with patch.dict(os.environ, {"JRPG_MODEL_TIMEOUT_SECONDS": value}):
                self.assertEqual(model_job.model_timeout_seconds(), expected)

    def test_gemini_and_openai_are_bounded_without_automatic_retries(self):
        for provider in ("gemini", "openai"):
            for failure in (False, True):
                client = Mock()
                client.models.generate_content.return_value = SimpleNamespace(text="ok")
                client.responses.create.return_value = SimpleNamespace(output_text="ok")
                call = client.models.generate_content if provider == "gemini" else client.responses.create
                if failure:
                    call.side_effect = TimeoutError("synthetic timeout")
                factory = Mock(return_value=client)
                types = SimpleNamespace(HttpOptions=lambda **kw: kw, HttpRetryOptions=lambda **kw: kw,
                                        GenerateContentConfig=lambda **kw: kw)
                modules = {"openai": SimpleNamespace(OpenAI=factory),
                           "google": SimpleNamespace(genai=SimpleNamespace(Client=factory)),
                           "google.genai": SimpleNamespace(types=types)}
                with patch.dict(sys.modules, modules), patch.object(example_sentence, "read_secret", return_value="fixture-not-a-key"), \
                     patch.object(example_sentence, "safety_settings", return_value=[]), \
                     patch.dict(os.environ, {"JRPG_MODEL_TIMEOUT_SECONDS": "75"}):
                    if failure:
                        with self.assertRaises(TimeoutError):
                            example_sentence.call_model(provider, "fixture", "fixture")
                    else:
                        self.assertEqual(example_sentence.call_model(provider, "fixture", "fixture"), "ok")
                kwargs = factory.call_args.kwargs
                if provider == "gemini":
                    self.assertEqual(kwargs["http_options"], {"timeout": 75000, "retry_options": {"attempts": 1}})
                else:
                    self.assertEqual(kwargs["timeout"], 75)
                    self.assertEqual(kwargs["max_retries"], 0)
                client.close.assert_called_once()
                call.assert_called_once()

    def test_wrapper_captures_unicode_rewrapped_streams_and_exit_codes(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            script = root / "fixture.py"
            out, err = root / "stdout.txt", root / "stderr.txt"
            for ending, expected in [("", 0), ("raise SystemExit(7)", 7), ("raise TimeoutError('fixture timeout')", 1)]:
                script.write_text("import sys, io\n"
                                  "sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')\n"
                                  "sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')\n"
                                  "print('日本語')\nprint('diagnostic', file=sys.stderr)\n" + ending, encoding="utf-8")
                result = subprocess.run([sys.executable, "-B", str(ROOT / "scripts/model_job.py"),
                                         "--stdout", str(out), "--stderr", str(err), str(script)],
                                        capture_output=True, timeout=10)
                self.assertEqual(result.returncode, expected, result.stderr)
                self.assertIn("日本語", out.read_text(encoding="utf-8"))
                self.assertIn("diagnostic", err.read_text(encoding="utf-8"))
                self.assertEqual(result.stdout, b"")

    def test_translation_status_is_structured_and_clears_on_success(self):
        tree = ast.parse((ROOT / "scripts/screenshot_translator.py").read_text(encoding="utf-8-sig"))
        fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "signal_ocr_completion")
        writes = {}
        namespace = {"os": os, "OVERLAY_DIR": "fixture", "OCR_DONE": "done",
                     "atomic_write_text": lambda path, text: writes.update({path: text}),
                     "request_completion_token": lambda: "request-123"}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), "fixture", "exec"), namespace)
        signal = namespace["signal_ocr_completion"]
        signal("504 DEADLINE_EXCEEDED")
        self.assertEqual(writes[os.path.join("fixture", "translation.status")], "request-123\n504 DEADLINE_EXCEEDED")
        self.assertEqual(writes["done"], "request-123")
        signal()
        self.assertEqual(writes[os.path.join("fixture", "translation.status")], "request-123\n")

    def test_translation_provider_failure_reaches_the_terminal_error_path(self):
        tree = ast.parse((ROOT / "scripts/screenshot_translator.py").read_text(encoding="utf-8-sig"))
        fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "translate_images")
        from typing import List, Tuple
        namespace = {"List": List, "Tuple": Tuple,
                     "call_translation_provider": Mock(side_effect=TimeoutError("fixture"))}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), "fixture", "exec"), namespace)
        with self.assertRaises(TimeoutError):
            namespace["translate_images"]([], [], [])


if __name__ == "__main__":
    unittest.main()
