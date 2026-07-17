# JRPG Translator Toolkit

JRPG Translator is a Windows toolkit for translating Japanese games while you
play. It combines screenshot translation, direct live-audio translation, and a
separate Japanese-learning explainer with customizable overlay windows.

The control panel works with a mouse and keyboard or entirely from a controller
through mapping tools such as JoyToKey, Steam Input, or DS4Windows.

## Features

- Screenshot translation with OpenAI and Google Gemini vision models.
- Near-live audio translation using OpenAI Realtime Translate or Gemini Live
  Translate models, without a separate transcription step.
- A dedicated Explainer for vocabulary, kanji readings, grammar, literal
  meaning, natural translations, nuance, and cultural context.
- Optional automatic saving of explanations as text files for later study.
- Independent Translator and Explainer overlays with configurable colors,
  fonts, borders, transparency, position, and size.
- Translation and explanation prompt profiles editable from the control panel.
- Shared OpenAI and Gemini model lists across translation and explanation.
- Selectable output language for live audio translation.
- JP-to-EN and EN-to-EN glossary profiles for names and terminology.
- Configurable hotkeys, spatial controller navigation, and optional dark mode.
- Non-activating overlays that can remain visible without taking focus from the
  game or pausing an emulator.

## Requirements

- Windows 10 or Windows 11.
- An OpenAI API key, a Gemini API key, or both.
- Internet access for translation and explanation requests.

The downloadable release includes a portable Python environment and compiled
AutoHotkey executables. A separate Python or AutoHotkey installation is not
needed when using the release package.

## Quick Start

1. Download and extract the
   [latest release](https://github.com/retrogamer0815/jrpg-translator-toolkit/releases/latest).
2. Run `JRPG Translator.exe`.
3. Add your API keys in the **API Keys** tab, or set `OPENAI_API_KEY` and/or
   `GEMINI_API_KEY` as Windows user environment variables.
4. Open the Translator overlay.
5. Choose a capture region or game window from
   **Screenshot Translation > Capture**.
6. Select the provider, model, prompt, and hotkeys you want to use.

The control panel saves settings automatically where appropriate. The **Save**
button becomes available when a manual save is needed.

## Translation Workflows

### Screenshot Translation

Use **Screenshot + Translate** for the fastest one-button workflow. You can also
take one or more screenshots first and translate them together, which is useful
when a sentence spans multiple dialogue boxes.

The screenshot prompt controls the requested output, so it can include a plain
translation, the original Japanese, kanji readings, speaker names, or any other
format useful for playing or studying.

### Live Audio Translation

In the **Audio Translation** tab:

1. Select the Windows playback device.
2. Choose OpenAI or Gemini and a compatible live translation model.
3. Select the output language.
4. Choose **Start Audio**, then enable or disable listening as needed.

Audio is streamed directly to the selected live translation model. Translated
lines appear at the bottom of the Translator overlay while older lines move
upward and remain available for scrolling.

### Japanese Explainer

The Explainer uses the most recent Japanese screenshot text. Choose
**Explain last jp. Text** or its configured hotkey to generate a separate
learning-focused explanation without replacing the translation.

Explanation prompts are independent from translation prompts, so they can be
tuned for a learner's level and preferred amount of detail.

When **Save explanations to textfiles** is enabled, generated explanations are
also stored in `Settings/Explanations` for use as study material.

## Controller Use

Map controller inputs to keyboard keys or JRPG Translator hotkeys with JoyToKey
or another controller mapper. In the control panel:

- Arrow keys move spatially between visible controls.
- Enter activates buttons, checkboxes, and dropdown selections.
- Page Up and Page Down switch tabs and work well when mapped to shoulder
  buttons.
- Mouse-wheel or arrow-key mappings can scroll the visible Translator or
  Explainer overlay even when it does not own game focus.

The overlays can be brought forward without becoming the active window. This
allows emulator options such as RetroArch's pause-when-inactive behavior to
remain enabled while translations are visible. Opening the control panel still
activates it normally.

## API Keys and Privacy

The recommended key-storage method is Windows user environment variables:

```text
OPENAI_API_KEY=your_key
GEMINI_API_KEY=your_key
```

The control panel can alternatively store keys in `Settings/.env`. This is a
plain-text file: do not commit it, upload it, or include it in shared archives.

Screenshots and audio sent for translation are processed by the selected API
provider. Review the provider's current data and privacy terms before use.

## Settings and Profiles

Portable settings live in the local `Settings` folder:

```text
Settings/
|-- control.ini
|-- .env                         # optional local API-key storage
|-- Screenshots/
|-- prompts/                     # screenshot translation prompts
|-- prompts_explain/             # explanation prompts
|-- glossaries/
|-- profiles/                    # Translator overlay profiles
`-- profiles_explainer/          # Explainer overlay profiles
```

Overlay size, position, appearance, and scrolling behavior are stored
independently for the Translator and Explainer.

The source repository and release include three prompt families for every
supported output language:

| Prompt pattern | Purpose |
| --- | --- |
| `Settings/prompts/default_<language>.txt` | Screenshot translation with a plain Japanese transcript |
| `Settings/prompts/default_with_kanji_reading_<language>.txt` | Screenshot translation with hiragana readings added to kanji words |
| `Settings/prompts_explain/default_<language>.txt` | Japanese-learning explanation in the selected language |

Available language labels are `en`, `de`, `fr`, `es`, `it`, `pt`, `nl`, `pl`,
`ru`, `uk`, `ko`, `zh-CN`, `zh-TW`, and `ja`. The screenshot prompts retain the
literal `Transcript:` and `Translation:` headings because the output parser uses
them; the translated content follows the language named by the profile.

Additional prompt profiles created through the control panel remain local and
are ignored by Git.

## Project Structure

| File | Purpose |
| --- | --- |
| `JRPG_Translator.ahk` | Main control panel and workflow orchestration |
| `bin/jrpg_overlay_C.ahk` | Translator and Explainer overlay windows |
| `bin/overlay.exe` | Compiled Translator and Explainer overlay used by release builds |
| `scripts/screenshot_translator.py` | Screenshot vision translation and output formatting |
| `scripts/live_audio_translator.py` | Direct streaming audio translation |
| `scripts/explainer.py` | Japanese-learning explanations |

Runtime messages and generated overlay text are exchanged through
`%TEMP%\JRPG_Overlay`.

## Running from Source

Install AutoHotkey v2 and run `JRPG_Translator.ahk`. The source version launches
`bin/jrpg_overlay_C.ahk`; compiled releases launch `bin/overlay.exe`.
The Python scripts require Python 3.12 and their listed dependencies, or the
portable Python environment included in a release package.

```powershell
py -3.12 -m pip install -r requirements.txt
```

Before sharing a build, verify that it does not contain `Settings/.env`, API
credentials, personal profiles, screenshots, logs, or other local settings.

## Troubleshooting

- If a request fails, verify the selected model name and confirm that the API
  key has access to that model.
- If the wrong playback source is translated, refresh and reselect the device
  in **Audio Translation**.
- If an overlay is missing, use the Open Translator or Open Explainer button and
  check its saved position on connected displays.
- If source files do not start, confirm that AutoHotkey v2 is being used rather
  than AutoHotkey v1.

## Credits

- [AutoHotkey](https://www.autohotkey.com/): GNU GPLv2.
- [Python](https://www.python.org/): PSF License.
- [PixelMplus](https://itouhiro.github.io/mplus-fonts/): SIL Open Font License
  1.1.
- Application icon by Miguel C Balandrano via Flaticon; attribution is required
  by the source license.

## License

The project source is released under the MIT License. See [LICENSE](LICENSE).
