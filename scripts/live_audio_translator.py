import asyncio
import base64
import configparser
import json
import os
import re
import sys
import tempfile
import time
import traceback
from pathlib import Path
from typing import List, Tuple

import numpy as np
import soundcard as sc
import websockets
from dotenv import load_dotenv


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent


def _load_dotenv_files():
    settings_dir = os.environ.get("SETTINGS_DIR", "").strip()
    candidates = []
    if settings_dir:
        candidates.append(Path(settings_dir) / ".env")
    candidates.extend([
        PROJECT_ROOT / "Settings" / ".env",
        PROJECT_ROOT / ".env",
    ])
    for path in candidates:
        if path.exists():
            load_dotenv(path, override=False, encoding="utf-8-sig")
            return


def _get_key(*names, file_var=None):
    bom = "\ufeff"
    for name in names:
        value = os.getenv(name) or os.getenv(bom + name, "")
        value = value.strip().strip('"').strip("'")
        if value:
            return value
    if file_var:
        path = (os.getenv(file_var) or os.getenv(bom + file_var, "")).strip().strip('"').strip("'")
        if path and os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    value = f.read().strip()
                if value:
                    return value
            except Exception:
                pass
    return ""


def load_glossary(path: str) -> List[Tuple[str, str]]:
    """Load source-to-target mappings using the translator's accepted formats."""
    entries: List[Tuple[str, str]] = []
    if not path or not os.path.isfile(path):
        return entries

    encodings = [
        "utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be",
        "cp1252", "cp932",
    ]
    text = None
    for encoding in encodings:
        try:
            with open(path, "r", encoding=encoding) as glossary_file:
                text = glossary_file.read()
            break
        except Exception:
            continue
    if text is None:
        return entries

    for raw_line in text.splitlines():
        line = raw_line.replace("\ufeff", "").strip()
        if not line or line.startswith("#"):
            continue
        source = target = None
        for separator in ("->", "→", "\t", ":", "="):
            if separator in line:
                source_part, target_part = line.split(separator, 1)
                source, target = source_part.strip(), target_part.strip()
                break
        if source and target:
            entries.append((source, target))
    return entries


def _sentence_case(text: str) -> str:
    lowered = text.lower()
    for index, char in enumerate(lowered):
        if char.isalpha():
            return lowered[:index] + char.upper() + lowered[index + 1:]
    return lowered


def _adapt_glossary_case(source: str, target: str, matched: str) -> str:
    """Adapt a replacement's case unless its source declares canonical casing."""
    source_first_letter = next((char for char in source if char.isalpha()), "")
    if source_first_letter.isupper():
        return target

    matched_letters = [char for char in matched if char.isalpha()]
    if not matched_letters:
        return target
    if all(char.isupper() for char in matched_letters):
        return target.upper()
    if all(char.islower() for char in matched_letters):
        return target.lower()

    if re.search(r"\s", source):
        words = re.findall(r"[^\W\d_]+", matched, flags=re.UNICODE)
        if words and all(
            word[0].isupper() and word[1:].islower() for word in words
        ):
            return target.title()

    if matched_letters[0].isupper() and all(
        char.islower() for char in matched_letters[1:]
    ):
        return _sentence_case(target)

    return target


def apply_target_glossary(
    text: str,
    glossary: List[Tuple[str, str]],
) -> str:
    """Apply case-insensitive, boundary-aware target-language replacements."""
    out = text
    for source, target in glossary:
        if re.search(r"\s", source):
            pattern = re.compile(re.escape(source), flags=re.IGNORECASE)
            out = pattern.sub(
                lambda match: _adapt_glossary_case(
                    source, target, match.group(0)
                ),
                out,
            )
        else:
            pattern = re.compile(
                rf"\b(?P<core>{re.escape(source)})(?P<suf>s|'s|’s)?\b",
                flags=re.IGNORECASE,
            )
            out = pattern.sub(
                lambda match: (
                    _adapt_glossary_case(source, target, match.group("core"))
                    + (match.group("suf") or "")
                ),
                out,
            )
    return out


def resolve_target_glossary_path(project_root: Path) -> str:
    """Resolve the TL-to-TL profile selected in the Terminology Overrides tab."""
    explicit_path = (
        os.environ.get("TL2TL_GLOSSARY_PATH", "").strip()
        or os.environ.get("EN2EN_GLOSSARY_PATH", "").strip()
    )
    if explicit_path:
        return explicit_path

    settings_dir = Path(
        os.environ.get("SETTINGS_DIR", "").strip()
        or project_root / "Settings"
    )
    profile = "default"
    for ini_name in ("control.ini", "config.ini"):
        ini_path = settings_dir / ini_name
        if not ini_path.is_file():
            continue
        for encoding in (
            "utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be",
            "cp1252", "cp932",
        ):
            try:
                config = configparser.ConfigParser(interpolation=None)
                with open(ini_path, "r", encoding=encoding) as ini_file:
                    config.read_file(ini_file)
                profile = (
                    config.get(
                        "cfg", "en2enGlossaryProfile", fallback="default"
                    ).strip()
                    or "default"
                )
                break
            except Exception:
                continue
        break

    return str(settings_dir / "glossaries" / profile / "en2en.txt")


_load_dotenv_files()

TL2TL_GLOSSARY_PATH = resolve_target_glossary_path(PROJECT_ROOT)
TL2TL_GLOSSARY = load_glossary(TL2TL_GLOSSARY_PATH)

OPENAI_API_KEY = _get_key(
    "OPENAI_API_KEY", "OPENAI_LOCAL_KEY", "OPENAI_API_KEY_LOCAL", "OPENAI_KEY",
    file_var="OPENAI_API_KEY_FILE",
)
GOOGLE_API_KEY = _get_key(
    "GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_LOCAL_KEY", "GOOGLE_LOCAL_KEY",
    file_var="GEMINI_API_KEY_FILE",
)

TEMP_DIR = os.environ.get("TEMP") or tempfile.gettempdir()
OVERLAY_DIR = os.path.join(TEMP_DIR, "JRPG_Overlay")
AUDIO_TXT = os.path.join(OVERLAY_DIR, "audio.txt")
LOG_TXT = os.path.join(OVERLAY_DIR, "audio_log.txt")
ERR_TXT = os.path.join(OVERLAY_DIR, "audio_error.txt")
os.makedirs(OVERLAY_DIR, exist_ok=True)

AUDIO_PROVIDER = (os.environ.get("AUDIO_PROVIDER", "openai") or "openai").strip().lower()
TRANSLATE_MODEL = (os.environ.get("TRANSLATE_MODEL", "gpt-realtime-translate") or "gpt-realtime-translate").strip()
GEMINI_AUDIO_MODEL = (
    os.environ.get("GEMINI_AUDIO_MODEL", "gemini-3.5-live-translate-preview")
    or "gemini-3.5-live-translate-preview"
).strip()
TARGET_LANGUAGE_CODE = (os.environ.get("TARGET_LANGUAGE_CODE", "en") or "en").strip()
TARGET_LANGUAGE_NAME = (os.environ.get("TARGET_LANGUAGE_NAME", "English") or "English").strip()
SPEAKER_NAME = os.environ.get("SPEAKER_NAME", "").strip()
DEBUG = os.environ.get("JRPG_DEBUG", "0").strip() == "1"

CAPTURE_RATE = 16000
BLOCK_DUR = 0.10
MAX_DISPLAY_CHARS = 1800


def _rotate_if_big(path, max_bytes=2_000_000, backups=3):
    try:
        if os.path.getsize(path) <= max_bytes:
            return
    except Exception:
        return
    for i in range(backups - 1, 0, -1):
        src = f"{path}.{i}"
        dst = f"{path}.{i + 1}"
        try:
            if os.path.exists(src):
                os.replace(src, dst)
        except Exception:
            pass
    try:
        os.replace(path, f"{path}.1")
    except Exception:
        pass


def log(line):
    if not DEBUG:
        return
    try:
        _rotate_if_big(LOG_TXT)
        with open(LOG_TXT, "a", encoding="utf-8", errors="ignore") as f:
            f.write(line.rstrip("\r\n") + "\n")
    except Exception:
        pass


def write_error(message):
    try:
        with open(ERR_TXT, "w", encoding="utf-8", errors="ignore") as f:
            f.write(message)
    except Exception:
        pass


def atomic_write_text(path, text):
    tmp = f"{path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)
    for _ in range(10):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.05)
        except FileNotFoundError:
            return
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)
    try:
        os.remove(tmp)
    except Exception:
        pass


def _pick_speaker():
    if SPEAKER_NAME:
        wanted = SPEAKER_NAME.lower()
        for speaker in sc.all_speakers():
            if wanted in speaker.name.lower():
                return speaker
    return sc.default_speaker()


def capture_blocks(speaker):
    block_size = int(CAPTURE_RATE * BLOCK_DUR)
    mic = sc.get_microphone(id=speaker.name, include_loopback=True)
    with mic.recorder(samplerate=CAPTURE_RATE, channels=2) as rec:
        while True:
            data = rec.record(numframes=block_size)
            if data.ndim == 2 and data.shape[1] > 1:
                mono = data.mean(axis=1)
            else:
                mono = data.reshape(-1)
            mono = np.clip(mono, -1.0, 1.0)
            pcm = (mono * 32767.0).astype("<i2", copy=False)
            yield pcm.tobytes()


def trim_display(text):
    text = text.strip()
    if len(text) <= MAX_DISPLAY_CHARS:
        return text
    return text[-MAX_DISPLAY_CHARS:].lstrip()


class TranscriptBuffer:
    def __init__(self):
        self.text = ""
        self.last_write = ""

    def append(self, delta):
        if not delta:
            return
        self.text = trim_display(self.text + delta)
        self.write()

    def replace(self, text):
        if text is None:
            return
        self.text = trim_display(text)
        self.write()

    def write(self):
        display_text = trim_display(
            apply_target_glossary(self.text, TL2TL_GLOSSARY)
        )
        if display_text != self.last_write:
            atomic_write_text(AUDIO_TXT, display_text)
            self.last_write = display_text


async def audio_sender(ws, speaker, make_event):
    blocks = capture_blocks(speaker)
    while True:
        chunk = await asyncio.to_thread(next, blocks)
        await ws.send(json.dumps(make_event(chunk)))


def _extract_text_parts(value):
    parts = []
    if isinstance(value, str):
        parts.append(value)
    elif isinstance(value, dict):
        for key in ("text", "transcript", "delta"):
            if isinstance(value.get(key), str):
                parts.append(value[key])
    elif isinstance(value, list):
        for item in value:
            parts.extend(_extract_text_parts(item))
    return parts


def _decode_ws_text(raw):
    if isinstance(raw, bytes):
        return raw.decode("utf-8", errors="replace")
    return raw


async def run_openai(speaker):
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY missing. Add it to Settings\\.env or the environment.")

    url = f"wss://api.openai.com/v1/realtime/translations?model={TRANSLATE_MODEL}"
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "OpenAI-Safety-Identifier": "jrpg-translator-local",
    }

    async with websockets.connect(url, additional_headers=headers, max_size=None) as ws:
        setup = {
            "type": "session.update",
            "session": {
                "audio": {
                    "output": {
                        "language": TARGET_LANGUAGE_CODE,
                    },
                },
            },
        }
        await ws.send(json.dumps(setup))

        def make_event(chunk):
            return {
                "type": "session.input_audio_buffer.append",
                "audio": base64.b64encode(chunk).decode("ascii"),
            }

        buf = TranscriptBuffer()
        sender = asyncio.create_task(audio_sender(ws, speaker, make_event))
        try:
            async for raw in ws:
                msg = json.loads(_decode_ws_text(raw))
                event_type = str(msg.get("type", ""))
                if DEBUG:
                    log(f"openai event: {event_type}")

                if event_type.endswith(".error") or event_type == "error":
                    log("openai error: " + json.dumps(msg, ensure_ascii=False))
                    continue

                if "output_transcript.delta" in event_type or "translation.delta" in event_type:
                    for key in ("delta", "text", "transcript"):
                        if isinstance(msg.get(key), str):
                            buf.append(msg[key])
                            break
                    continue

                if "output_transcript.done" in event_type or "translation.done" in event_type:
                    for key in ("transcript", "text"):
                        if isinstance(msg.get(key), str):
                            buf.replace(msg[key])
                            break
                    continue

                for key in ("output_transcript", "translation", "response"):
                    for text in _extract_text_parts(msg.get(key)):
                        buf.append(text)
        finally:
            sender.cancel()


async def run_gemini(speaker):
    if not GOOGLE_API_KEY:
        raise RuntimeError("GEMINI_API_KEY or GOOGLE_API_KEY missing. Add it to Settings\\.env or the environment.")

    model = GEMINI_AUDIO_MODEL
    if not model.startswith("models/"):
        model = "models/" + model

    url = (
        "wss://generativelanguage.googleapis.com/ws/"
        "google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent"
        f"?key={GOOGLE_API_KEY}"
    )

    async with websockets.connect(url, max_size=None) as ws:
        setup = {
            "setup": {
                "model": model,
                "generationConfig": {
                    "responseModalities": ["AUDIO"],
                    "translationConfig": {
                        "targetLanguageCode": TARGET_LANGUAGE_CODE,
                        "echoTargetLanguage": True,
                    },
                },
                "inputAudioTranscription": {},
                "outputAudioTranscription": {},
            }
        }
        await ws.send(json.dumps(setup))
        setup_ack = _decode_ws_text(await ws.recv())
        log("gemini setup: " + setup_ack[:500])

        def make_event(chunk):
            return {
                "realtimeInput": {
                    "audio": {
                        "data": base64.b64encode(chunk).decode("ascii"),
                        "mimeType": f"audio/pcm;rate={CAPTURE_RATE}",
                    }
                }
            }

        buf = TranscriptBuffer()
        sender = asyncio.create_task(audio_sender(ws, speaker, make_event))
        try:
            async for raw in ws:
                msg = json.loads(_decode_ws_text(raw))
                server = msg.get("serverContent") or msg.get("server_content") or {}
                out_tr = server.get("outputTranscription") or server.get("output_transcription")
                for text in _extract_text_parts(out_tr):
                    buf.append(text)

                # Some SDK/proxy versions surface transcript chunks inside modelTurn parts.
                turn = server.get("modelTurn") or server.get("model_turn") or {}
                for part in turn.get("parts", []) or []:
                    for text in _extract_text_parts(part):
                        buf.append(text)
        finally:
            sender.cancel()


async def main_async():
    try:
        atomic_write_text(AUDIO_TXT, "")
    except Exception:
        pass

    speaker = _pick_speaker()
    if speaker is None:
        raise RuntimeError("No playback device found. Check Windows sound settings.")

    print("=== JRPG live audio translation ===")
    print(f"Provider : {AUDIO_PROVIDER}")
    print(f"Speaker  : {speaker.name}")
    print(f"Overlay  : {AUDIO_TXT}")
    if AUDIO_PROVIDER == "gemini":
        print(f"Model    : {GEMINI_AUDIO_MODEL}")
        await run_gemini(speaker)
    else:
        print(f"Model    : {TRANSLATE_MODEL}")
        await run_openai(speaker)


def main():
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        pass
    except Exception:
        detail = traceback.format_exc()
        write_error(detail)
        try:
            atomic_write_text(AUDIO_TXT, "Live audio error - see audio_error.txt")
        except Exception:
            pass
        print(detail, file=sys.stderr)
        time.sleep(0.5)
        raise


if __name__ == "__main__":
    if "--list-speakers" in sys.argv:
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
        except Exception:
            pass
        for speaker in sc.all_speakers():
            print(speaker.name)
        try:
            sys.stdout.flush()
        except Exception:
            pass
        os._exit(0)
    main()
