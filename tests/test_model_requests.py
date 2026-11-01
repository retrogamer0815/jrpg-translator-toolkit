"""Offline request deadlines, diagnostic transport and provider-client contracts."""
import ast
import io
import os
from pathlib import Path
import re
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
        for value, expected in [("75", 75), ("invalid", 75), ("900", 90), ("0", 15),
                                ("nan", 75), ("inf", 75), ("1e999", 75)]:
            with patch.dict(os.environ, {"JRPG_MODEL_TIMEOUT_SECONDS": value}):
                self.assertEqual(model_job.model_timeout_seconds(), expected)

    def test_study_deadlines_are_not_shortened_by_inherited_translation_timeout(self):
        for legacy in ("15", "75", "invalid"):
            with patch.dict(os.environ, {"JRPG_MODEL_TIMEOUT_SECONDS": legacy}):
                for kind, seconds in [("explanation", 300), ("evaluation", 600), ("example", 120)]:
                    self.assertEqual(model_job.model_timeout_seconds(kind), seconds)
        with self.assertRaises(KeyError):
            model_job.model_timeout_seconds("misspelled-task")

    def test_ui_deadlines_match_provider_defaults_and_all_generation_entry_points(self):
        source = (ROOT / "JRPG Translator.ahk").read_text(encoding="utf-8-sig")
        policy = re.search(r"(?ms)^CPModelRequestTimeoutMs\(.*?^}", source).group()
        defaults = dict((kind, float(seconds)) for kind, seconds in
                        re.findall(r'"(translation|explanation|quick_explanation|evaluation|example)", (\d+)', policy))
        self.assertEqual(defaults, model_job.MODEL_TIMEOUT_SECONDS)
        self.assertIn('(seconds[requestKind] + 45) * 1000', policy)
        for function, kind in [("StudyReaderGenerateNewVersion", "explanation"),
                               ("StudyReaderGenerateVocabularyExample", "example"),
                               ("StudyCandidatesRunBridge", "evaluation")]:
            body = re.search(r"(?ms)^" + function + r"\(.*?^}", source).group()
            self.assertIn('CPModelRequestTimeoutMs("' + kind + '")', body)
        explain = re.search(r"(?ms)^ExplainNowCore\(.*?^}", source).group()
        self.assertIn('CPModelRequestTimeoutMs(quick ? "quick_explanation" : "explanation")', explain)
        self.assertNotIn('EnvSet("JRPG_MODEL_TIMEOUT_SECONDS"', explain)
        self.assertNotIn('>= 120000', explain)
        overlay = (ROOT / "bin/overlay.ahk").read_text(encoding="utf-8-sig")
        self.assertIn('__TranslationTimeoutMs := 120000', overlay)
        self.assertIn('EnvSet("JRPG_MODEL_TIMEOUT_SECONDS", "75")', overlay)

    def test_explainer_uses_explanation_policy_for_both_providers(self):
        source = (ROOT / "scripts/explainer.py").read_text(encoding="utf-8-sig")
        tree = ast.parse(source)
        assignment = next(n for n in tree.body if isinstance(n, ast.Assign)
                          and any(isinstance(t, ast.Name) and t.id == "MODEL_TIMEOUT_SECONDS" for t in n.targets))
        namespace = {"model_timeout_seconds": model_job.model_timeout_seconds, "QUICK_EXPLANATION": False}
        with patch.dict(os.environ, {"JRPG_MODEL_TIMEOUT_SECONDS": "75"}):
            exec(compile(ast.Module(body=[assignment], type_ignores=[]), "fixture", "exec"), namespace)
        self.assertEqual(namespace["MODEL_TIMEOUT_SECONDS"], 300)
        self.assertIn('timeout=int(MODEL_TIMEOUT_SECONDS * 1000)', source)
        self.assertIn('timeout=MODEL_TIMEOUT_SECONDS,', source)

    def test_gemini_and_openai_are_bounded_without_automatic_retries(self):
        for provider in ("gemini", "openai"):
            for kind, failure in ((kind, failure) for kind in model_job.MODEL_TIMEOUT_SECONDS
                                  for failure in (False, True)):
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
                            example_sentence.call_model(provider, "fixture", "fixture", request_kind=kind)
                    else:
                        self.assertEqual(example_sentence.call_model(provider, "fixture", "fixture", request_kind=kind), "ok")
                kwargs = factory.call_args.kwargs
                expected = model_job.MODEL_TIMEOUT_SECONDS[kind]
                if provider == "gemini":
                    self.assertEqual(kwargs["http_options"], {"timeout": expected * 1000, "retry_options": {"attempts": 1}})
                else:
                    self.assertEqual(kwargs["timeout"], expected)
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
