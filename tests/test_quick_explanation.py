"""Run the real explainer entry point with fake SDKs and isolated output paths."""
from contextlib import closing
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]

DRIVER = r'''
import json, os, runpy, sys, types, socket
from pathlib import Path
from types import SimpleNamespace as NS
# No provider calls or personal .env loading are possible in this fixture.
socket.socket = lambda *a, **k: (_ for _ in ()).throw(AssertionError("Network forbidden"))
sys.modules["dotenv"] = NS(load_dotenv=lambda *a, **k: None, find_dotenv=lambda **k: "")
def record(**kw):
    Path("request.json").write_text(json.dumps(kw), encoding="utf-8")
    if os.environ.get("FIXTURE_FAIL"):
        raise TimeoutError("synthetic request timed out")
    return NS(output_text="Vocabulary:\n言葉（ことば）: word\n\nKey grammar:\nが marks the subject.",
              text="Vocabulary:\n言葉（ことば）: word\n\nKey grammar:\nが marks the subject.")
def client(**kw):
    Path("client.json").write_text(json.dumps(kw), encoding="utf-8")
    return NS(responses=NS(create=record), models=NS(generate_content=record))
sys.modules["openai"] = NS(OpenAI=client)
enum = NS(HARM_CATEGORY_HARASSMENT=1, HARM_CATEGORY_HATE_SPEECH=2,
          HARM_CATEGORY_SEXUALLY_EXPLICIT=3, HARM_CATEGORY_DANGEROUS_CONTENT=4, OFF=0)
types_stub = NS(HttpOptions=lambda **k:k, HttpRetryOptions=lambda **k:k,
                GenerateContentConfig=lambda **k:k, SafetySetting=lambda **k:k,
                HarmCategory=enum, HarmBlockThreshold=enum,
                Part=NS(from_text=lambda **k:k, from_bytes=lambda data, mime_type: {"data":data.hex(), "mime_type":mime_type}),
                Content=lambda **k:k)
sys.modules["google"] = NS(genai=NS(Client=client))
sys.modules["google.genai"] = NS(types=types_stub)
script = sys.argv.pop(1)
runpy.run_path(script, run_name="__main__")
'''


class QuickExplanation(unittest.TestCase):
    def run_explainer(self, root, *, quick, provider="openai", fail=False,
                      source_paths_file=None, source_text_file=None, include_screenshots=False):
        (root / "driver.py").write_text(DRIVER, encoding="utf-8")
        (root / "source.txt").write_text("言葉が分かる。", encoding="utf-8")
        (root / "paths.txt").write_text("", encoding="utf-8")
        env = {key: value for key, value in os.environ.items()
               if key.upper() in {"SYSTEMROOT", "WINDIR", "PATH"}}
        env.update(TEMP=str(root), TMP=str(root), SETTINGS_DIR=str(root / "settings"),
                   EXPLAIN_SOURCE_TEXT_FILE=str(source_text_file or root / "source.txt"),
                   EXPLAIN_SOURCE_PATHS_FILE=str(source_paths_file or root / "paths.txt"),
                   EXPLAIN_PROVIDER=provider, OPENAI_API_KEY="fixture", GEMINI_API_KEY="fixture",
                   EXPLAIN_MODEL="detailed-fixture" if not quick else "quick-fixture",
                   GEMINI_EXPLAIN_MODEL="detailed-fixture" if not quick else "quick-fixture",
                   EXPLAIN_PROMPT_FILE=str(ROOT / "Settings/prompts_quick_explain/default_en.txt") if quick else "",
                   SAVE_EXPLAINS="1", SAVE_STUDY_LIBRARY="1", STUDY_LIBRARY_SCREENSHOTS="1" if include_screenshots else "0",
                   EXPLAIN_ERROR_FILE=str(root / "error.txt"), JRPG_REQUEST_ID="quick" if quick else "detailed",
                   JRPG_MODEL_TIMEOUT_SECONDS="15", FIXTURE_FAIL="1" if fail else "")
        result = subprocess.run([sys.executable, "-B", str(root / "driver.py"),
                                 str(ROOT / "scripts/explainer.py"), *(["--quick-explanation"] if quick else [])],
                                env=env, cwd=root, capture_output=True, timeout=15)
        self.assertEqual(result.returncode, 1 if fail else 0, result.stderr.decode("utf-8", errors="replace"))
        client = json.loads((root / "client.json").read_text(encoding="utf-8"))
        request = json.loads((root / "request.json").read_text(encoding="utf-8"))
        return client, request

    def image_fixture(self, root):
        from PIL import Image
        images = [root / "日本語-first.png", root / "second.png"]
        for path, color in zip(images, ("red", "green")):
            Image.new("RGB", (8, 8), color).save(path)
        return images

    def assert_archived_images(self, root, expected):
        library = root / "settings/Study Library"
        with closing(sqlite3.connect(library / "study_library.db")) as db:
            rows = db.execute("SELECT relative_path, sort_order FROM media ORDER BY sort_order").fetchall()
            self.assertEqual(len(rows), len(expected), "Every capture, including the first, must reach the Library")
            self.assertEqual([row[1] for row in rows], list(range(1, len(expected) + 1)))
            for (relative, _), original in zip(rows, expected):
                self.assertEqual((library / relative).read_bytes(), original)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM media m JOIN explanations e ON e.id=m.explanation_id").fetchone()[0], len(expected))
        return [library / relative for relative, _ in rows]

    def test_bom_prefixed_paths_keep_first_image_for_both_providers(self):
        for provider in ("openai", "gemini"):
            for count in (1, 2):
                with self.subTest(provider=provider, count=count), tempfile.TemporaryDirectory() as temp:
                    root = Path(temp)
                    images = self.image_fixture(root)[:count]
                    manifest = root / "bom-paths.txt"
                    manifest.write_text("\r\n".join(map(str, images)), encoding="utf-8-sig")
                    japanese = root / "bom-source.txt"
                    japanese.write_text("言葉が分かる。", encoding="utf-8-sig")
                    self.run_explainer(root, quick=True, provider=provider, source_paths_file=manifest, include_screenshots=True)
                    self.assertFalse((root / "settings").exists(), "Quick never archives images or text")
                    _, request = self.run_explainer(root, quick=False, provider=provider, source_paths_file=manifest,
                                                    source_text_file=japanese, include_screenshots=True)
                    parts = request["input"][0]["content"] if provider == "openai" else request["contents"][0]["parts"]
                    self.assertEqual(len(parts) - 1, count, "Every source image also reaches the model")
                    self.assertNotIn("\ufeff", parts[0]["text"])
                    self.assert_archived_images(root, [p.read_bytes() for p in images])

    @unittest.skipUnless(Path(r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe").is_file(), "AutoHotkey v2 required")
    def test_real_ahk_snapshot_survives_cleanup_in_library_and_anki(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            images = self.image_fixture(root)
            expected = [p.read_bytes() for p in images]
            (root / "last_src.txt").write_text("\r\n".join(map(str, images)), encoding="utf-8")
            (root / "last_jp.txt").write_text("言葉が分かる。", encoding="utf-8")
            source = (ROOT / "JRPG Translator.ahk").read_text(encoding="utf-8-sig")
            function = re.search(r"(?ms)^CPExplanationSnapshot\(.*?^}", source).group()
            # Only redirect the temporary root; all snapshot writing is production code.
            function = function.replace('A_Temp "\\JRPG_ExplanationRequests', 'A_ScriptDir "\\JRPG_ExplanationRequests')
            script = root / "snapshot.ahk"
            script.write_text('#Requires AutoHotkey v2.0\n#SingleInstance Off\n'
                              'global overlayDir := A_ScriptDir\n'
                              'snapshot := CPExplanationSnapshot("fixture")\n'
                              'FileAppend(snapshot, "*", "UTF-8-RAW")\nExitApp()\n' + function, encoding="utf-8-sig")
            result = subprocess.run([r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe", "/ErrorStdOut", str(script)],
                                    capture_output=True, timeout=10, creationflags=subprocess.CREATE_NO_WINDOW)
            self.assertEqual(result.returncode, 0, result.stdout.decode("utf-8", errors="replace"))
            snapshot = Path(result.stdout.decode("utf-8").strip())
            self.assertEqual(snapshot.parent.resolve(), (root / "JRPG_ExplanationRequests").resolve())
            self.assertEqual((snapshot / "source.txt").read_text(encoding="utf-8"), "言葉が分かる。")
            self.run_explainer(root, quick=False, source_paths_file=snapshot / "paths.txt",
                               source_text_file=snapshot / "source.txt", include_screenshots=True)
            # Matches the control panel's cleanup when the worker exits.
            shutil.rmtree(snapshot)
            media = self.assert_archived_images(root, expected)
            sys.path.insert(0, str(ROOT / "scripts"))
            from study_library import detail
            detail_output = root / "reader-detail"
            detail_output.mkdir()
            self.assertEqual(detail(root / "settings/Study Library/study_library.db", detail_output, 1, 1), 0)
            rows = (detail_output / "media.tsv").read_text(encoding="utf-8-sig").splitlines()
            self.assertEqual(len(rows), len(images))
            self.assertTrue(all(row.split("\t")[-1] == "1" for row in rows), "Study Reader sees all archived images as available")
            from anki_bridge import prepare_anki_screenshot
            for screenshot in media:
                filename, encoded, size, dimensions = prepare_anki_screenshot(screenshot)
                self.assertTrue(filename and encoded)
                self.assertGreater(size, 0)
                self.assertEqual(dimensions, (8, 8))

    def test_quick_then_detailed_keep_prompt_timeout_and_archives_independent(self):
        for provider in ("openai", "gemini"):
            with self.subTest(provider=provider), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                client, request = self.run_explainer(root, quick=True, provider=provider)
                self.assertEqual(client.get("timeout", client.get("http_options", {}).get("timeout")),
                                 120 if provider == "openai" else 120000)
                prompt = request.get("input", request.get("contents"))
                self.assertIn("言葉が分かる。", prompt)
                self.assertIn("Vocabulary", prompt)
                self.assertNotIn("JRPG_TRANSLATOR_METADATA", prompt)
                self.assertFalse((root / "settings").exists(), "Quick must not create either archive, even with both flags set")
                overlay = root / "JRPG_Overlay"
                self.assertEqual((overlay / "explainer.done").read_text(), "quick")
                self.assertIn("Vocabulary", (overlay / "explainer.txt").read_text(encoding="utf-8"))
                client, request = self.run_explainer(root, quick=False, provider=provider)
                self.assertEqual(client.get("timeout", client.get("http_options", {}).get("timeout")),
                                 300 if provider == "openai" else 300000)
                self.assertEqual(request["model"], "detailed-fixture")
                self.assertIn("JRPG_TRANSLATOR_METADATA", request.get("input", request.get("contents")))
                self.assertTrue((root / "settings/Study Library/study_library.db").is_file())
                self.assertTrue(list((root / "settings/Explanations").rglob("*.txt")))
                self.assertEqual((overlay / "explainer.done").read_text(), "detailed")

    def test_quick_failure_then_retry_does_not_archive_or_lose_error_status(self):
        for provider in ("openai", "gemini"):
            with self.subTest(provider=provider), tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                self.run_explainer(root, quick=True, provider=provider, fail=True)
                self.assertIn("timed out", (root / "error.txt").read_text(encoding="utf-8"))
                self.assertEqual((root / "JRPG_Overlay/explainer.done").read_text(), "quick")
                self.assertFalse((root / "settings").exists())
                (root / "error.txt").unlink()
                self.run_explainer(root, quick=True, provider=provider)
                self.assertFalse((root / "error.txt").exists())
                self.assertFalse((root / "settings").exists())


if __name__ == "__main__":
    unittest.main()
