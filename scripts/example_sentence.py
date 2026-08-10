#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Generate one structured learner-friendly example sentence.

The AutoHotkey front end passes source text through UTF-8 files and receives a
tab-separated row whose text fields are UTF-8 hex.  This keeps arbitrary
Japanese and translated text out of command-line quoting.
"""

import json
import os
import re
import sys
import warnings
from pathlib import Path

from dotenv import find_dotenv, load_dotenv

# This packaged runtime warning is informational and unrelated to generation
# failures.  Suppress it here so a genuine API/model error remains readable.
warnings.filterwarnings(
    "ignore",
    message=r"You are using cryptography on a 32-bit Python.*",
    category=UserWarning,
    module=r"cryptography\.hazmat\.backends\.openssl\.backend",
)


def read_text(path: str) -> str:
    if not path:
        return ""
    try:
        return Path(path).read_text(encoding="utf-8-sig").strip()
    except Exception:
        return ""


def load_project_environment() -> None:
    root = Path(__file__).resolve().parents[1]
    settings_dir = os.environ.get("SETTINGS_DIR", "").strip()
    candidates = []
    if settings_dir:
        candidates.append(Path(settings_dir) / ".env")
    candidates.extend((root / "Settings" / ".env", root / ".env"))
    for candidate in candidates:
        if candidate.is_file():
            load_dotenv(candidate, override=False, encoding="utf-8-sig")
            return
    located = find_dotenv(usecwd=True)
    if located:
        load_dotenv(located, override=False, encoding="utf-8-sig")


def clean_secret(value: str) -> str:
    return (value or "").strip().strip('"').strip("'")


def read_secret(names, file_name: str) -> str:
    for name in names:
        value = clean_secret(os.environ.get(name, ""))
        if value:
            return value
    key_file = clean_secret(os.environ.get(file_name, ""))
    if key_file and Path(key_file).is_file():
        return clean_secret(read_text(key_file))
    return ""


def safety_settings(types):
    return [
        types.SafetySetting(
            category=category,
            threshold=types.HarmBlockThreshold.OFF,
        )
        for category in (
            types.HarmCategory.HARM_CATEGORY_HARASSMENT,
            types.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            types.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            types.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
        )
    ]


def build_prompt(term: str, definition: str, context: str) -> str:
    context = context[:12000]
    return f"""You are helping a Japanese-language learner create one Anki example.

Target vocabulary:
{term}

Vocabulary definition and notes (its language determines the translation language):
{definition}

Source explanation for context only:
{context}

Create exactly one natural, memorable example sentence that uses the target
vocabulary (a natural inflected form is allowed).  Aim at upper-beginner to
intermediate level: useful and vivid, but neither childish nor needlessly
complex.  Do not copy the source-game sentence.

Return only one JSON object with these exact keys:
{{
  "plain_japanese": "the sentence in normal Japanese without reading assistance",
  "reading_japanese": "the identical sentence with kana readings immediately after kanji as 漢字(かんじ)",
  "translation": "a natural translation in the same language used by the vocabulary definition above"
}}

All three values must be single lines.  Do not add Markdown, labels, commentary,
romaji, or alternative sentences.  The first two values must express exactly
the same sentence; only the reading assistance may differ.
"""


def extract_json_object(raw: str):
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)

    # raw_decode deliberately returns where the first JSON value ended.  That
    # allows a correct object to survive occasional trailing model commentary
    # or even an accidental second object, while the object itself must still
    # be fully valid JSON.
    decoder = json.JSONDecoder()
    candidate_starts = [0]
    candidate_starts.extend(
        index for index, character in enumerate(text) if character == "{"
    )
    seen = set()
    for start in candidate_starts:
        if start in seen:
            continue
        seen.add(start)
        candidate = text[start:].lstrip()
        if not candidate:
            continue
        try:
            value, _end = decoder.raw_decode(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise ValueError("The model did not return the requested JSON object.")


def single_line(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_result(value):
    plain = single_line(value.get("plain_japanese"))
    reading = single_line(value.get("reading_japanese"))
    translation = single_line(value.get("translation"))
    # Be defensive when a model ignores the plain-Japanese requirement.
    plain = re.sub(
        r"([ぁ-ゖァ-ヺ一-龯々〆ヵヶー]+)[(（][ぁ-ゖァ-ヺー・･ /]+[)）]",
        r"\1",
        plain,
    )
    if not plain or not reading or not translation:
        raise ValueError("The model returned an incomplete example sentence.")
    if not re.search(r"[ぁ-ゖァ-ヺ一-龯々〆ヵヶ]", plain):
        raise ValueError("The generated example does not contain Japanese text.")
    return plain, reading, translation


def call_model(provider: str, model: str, prompt: str) -> str:
    if provider == "gemini":
        try:
            from google import genai
            from google.genai import types
        except Exception as exc:
            raise RuntimeError("The google-genai package could not be loaded.") from exc
        key = read_secret(
            ("GEMINI_API_KEY", "GOOGLE_API_KEY", "GEMINI_LOCAL_KEY", "GOOGLE_LOCAL_KEY"),
            "GEMINI_API_KEY_FILE",
        )
        if not key:
            raise RuntimeError("Gemini API key missing.")
        if model.startswith("models/"):
            model = model[7:]
        # Keep the Client itself alive for the complete synchronous request.
        # Accessing .models from a temporary Client can leave its HTTP session
        # closed before generate_content begins on some google-genai versions.
        client = genai.Client(api_key=key)
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.7,
                    response_mime_type="application/json",
                    safety_settings=safety_settings(types),
                ),
            )
        finally:
            try:
                client.close()
            except Exception:
                pass
        return (getattr(response, "text", "") or "").strip()

    if provider == "openai":
        try:
            from openai import OpenAI
        except Exception as exc:
            raise RuntimeError("The OpenAI package could not be loaded.") from exc
        bom = "\ufeff"
        key = read_secret(
            (
                "OPENAI_API_KEY", bom + "OPENAI_API_KEY", "OPENAI_LOCAL_KEY",
                bom + "OPENAI_LOCAL_KEY", "OPENAI_API_KEY_LOCAL",
                bom + "OPENAI_API_KEY_LOCAL", "OPENAI_KEY", bom + "OPENAI_KEY",
            ),
            "OPENAI_API_KEY_FILE",
        )
        if not key:
            raise RuntimeError("OpenAI API key missing.")
        client = OpenAI(api_key=key)
        try:
            response = client.responses.create(model=model, input=prompt)
        finally:
            try:
                client.close()
            except Exception:
                pass
        return (getattr(response, "output_text", "") or "").strip()

    raise RuntimeError(f"Unknown provider: {provider}")


def text_hex(value: str) -> str:
    return value.encode("utf-8").hex()


def write_result(path: str, plain: str, reading: str, translation: str) -> None:
    Path(path).write_text(
        "\t".join(("ok", text_hex(plain), text_hex(reading), text_hex(translation)))
        + "\n",
        encoding="utf-8",
    )


def main() -> int:
    load_project_environment()
    provider = os.environ.get("EXPLAIN_PROVIDER", "openai").strip().lower()
    if provider == "gemini":
        model = os.environ.get("GEMINI_EXPLAIN_MODEL", "gemini-2.5-flash").strip()
    else:
        model = os.environ.get("EXPLAIN_MODEL", "gpt-4o-mini").strip()
    term = read_text(os.environ.get("EXAMPLE_TERM_FILE", ""))
    definition = read_text(os.environ.get("EXAMPLE_DEFINITION_FILE", ""))
    context = read_text(os.environ.get("EXAMPLE_CONTEXT_FILE", ""))
    result_path = os.environ.get("EXAMPLE_RESULT_FILE", "").strip()
    if not term or not definition or not result_path:
        raise RuntimeError("The vocabulary term, definition, or result path is missing.")
    raw = call_model(provider, model, build_prompt(term, definition, context))
    plain, reading, translation = normalize_result(extract_json_object(raw))
    write_result(result_path, plain, reading, translation)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Example generation failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
