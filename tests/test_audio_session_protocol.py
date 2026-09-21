import ast
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch


SOURCE = Path(__file__).resolve().parents[1] / "scripts" / "live_audio_translator.py"
TREE = ast.parse(SOURCE.read_text(encoding="utf-8-sig"), filename=str(SOURCE))
FUNCTION_NAMES = {"atomic_write_text", "audio_session_identity", "claim_audio_session", "release_audio_session"}
NODES = [
    node
    for node in TREE.body
    if isinstance(node, ast.FunctionDef) and node.name in FUNCTION_NAMES
]
CODE = compile(ast.Module(body=NODES, type_ignores=[]), str(SOURCE), "exec")


class AudioSessionProtocolTests(unittest.TestCase):
    def namespace(self, marker: Path):
        namespace = {
            "os": os,
            "time": time,
            "Path": Path,
            "AUDIO_SESSION_FILE": str(marker),
        }
        exec(CODE, namespace)
        return namespace

    def test_worker_claims_and_releases_its_process_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "nested" / "audio-session.pid"
            namespace = self.namespace(marker)
            namespace["claim_audio_session"]()
            identity = marker.read_text(encoding="utf-8")
            self.assertEqual(identity, namespace["audio_session_identity"]())
            self.assertRegex(identity, rf"^JRPG_AUDIO_V2\t{os.getpid()}\t[0-9A-F]{{16}}$")
            namespace["release_audio_session"]()
            self.assertFalse(marker.exists())

    def test_worker_does_not_remove_newer_owner_marker(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "audio-session.pid"
            namespace = self.namespace(marker)
            namespace["claim_audio_session"]()
            marker.write_text(str(os.getpid() + 1), encoding="utf-8")
            namespace["release_audio_session"]()
            self.assertTrue(marker.exists())

    def test_recycled_pid_marker_is_not_removed(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "audio-session.pid"
            namespace = self.namespace(marker)
            marker.write_text(f"JRPG_AUDIO_V2\t{os.getpid()}\t0000000000000000", encoding="utf-8")
            namespace["release_audio_session"]()
            self.assertTrue(marker.exists())

    def test_identity_failure_does_not_publish_a_pid_only_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "audio-session.pid"
            namespace = self.namespace(marker)
            with patch.dict(namespace, {"audio_session_identity": lambda: (_ for _ in ()).throw(OSError("denied"))}):
                namespace["claim_audio_session"]()
            self.assertFalse(marker.exists())

    def test_tracking_failure_does_not_block_audio_startup(self):
        with tempfile.TemporaryDirectory() as directory:
            marker = Path(directory) / "audio-session.pid"
            namespace = self.namespace(marker)
            with patch.object(namespace["Path"], "mkdir", side_effect=OSError("denied")):
                namespace["claim_audio_session"]()
            self.assertFalse(marker.exists())


if __name__ == "__main__":
    unittest.main()
