## 🕹️ JRPG Translator

**JRPG Translator** is a live on-screen translation toolkit for Japanese retro and modern games.
It provides **screen**, **audio**, and **contextual explanation** overlays — ideal for learning Japanese while playing, or simply enjoying untranslated games.

The system combines **AutoHotkey GUIs** and **Python back-end scripts**, with support for **OpenAI**, **Google Gemini**, and **local faster-whisper** speech recognition.

---

### ✨ Features

* 🎮 **Fully Hotkey-Driven**
  Every function — from screenshot capture and translation to toggling overlays or starting audio transcription — can be mapped to custom hotkeys.
  Works great with controller mapping tools like *JoyToKey*, *Steam Input*, or *DS4Windows* for seamless in-game control.

* 🧠 **Prompt & Model Flexibility**
  Change or edit system prompts, try different OpenAI / Gemini models, and even add new providers directly through the UI — without restarting the app or touching config files.

* 🖼️ **Screenshot Translation**
  Translate in-game dialogue boxes or menus instantly. Capture any region or window and see the result appear in the overlay.

* 🎙️ **Audio Translation**
  Translate Japanese voice acting with near real time text output by using OpenAIs and Geminis realtime/live audio models.

* 💬 **Grammar Explainer**
  Turn any captured line into an English breakdown of grammar, vocabulary, and nuance — perfect for Japanese learners. Save these optional as text files automatically to use them as study material.

* 🧩 **Dual Glossary System**
  Custom JP→EN and EN→EN glossaries ensure consistent terminology and phrasing across translations.

* 🪟 **JRPG-Style Overlays**
  Clean, resizable floating windows that mimic RPG text boxes — with customizable fonts, borders, and transparency.

* 🧰 **All Settings in One Place**
  The Control Panel manages everything — paths, models, colors, hotkeys, and provider settings — with changes applied instantly.

---

### 🏗️ How It Works

The project has three main layers:

| Component                                                                      | Purpose                                                                                                  | Language   |
| ------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------- | ---------- |
| `JRPG_Control_Panel.ahk`                                                       | Main control panel – lets you start/stop overlays, manage API keys, choose models, and set color themes. | AutoHotkey |
| `jrpg_overlay.ahk`                                                             | On-screen “JRPG-style” overlay window that displays live translations or explanations.                   | AutoHotkey |
| `scripts/` (`screenshot_translator.py`, `audio_translator.py`, `explainer.py`) | Python back-ends handling OCR, ASR, translation, and prompt processing.                                  | Python 3   |

The AHK control panel sets environment variables and launches the relevant Python modules:

* `screenshot_translator.py` → translates captured screenshots using AI vision models
* `audio_translator.py` → transcribes and translates live audio
* `explainer.py` → generates grammar and vocabulary breakdowns for the last translated Japanese line

All output is written to a shared temporary folder (`%TEMP%\JRPG_Overlay\`), where the overlays read it in real time.

---

### 🧩 Requirements

* **Windows 10 / 11**
* **Python 3.12.** (portable environment included in the release)
* OpenAI and/or Gemini API keys
* Internet connection (for OpenAI / Gemini modes)

---

### 🪄 Quick Start

1. **Download** or clone the repository

   ```
   git clone https://github.com/<yourname>/JRPG-Translator.git
   ```
2. **Run** `JRPG Translator.exe` (or `.ahk` if using AutoHotkey v2)
3. **Enter your API keys** in the API Keys tab or in Windows under System Environment Variables.
 
   OPENAI_API_KEY="your_openai_key"
   GEMINI_API_KEY="your_google_key"
   ```
4. **Choose provider and models** under the “Audio” and “Screenshot” tabs.
5. **Start overlays** for Translator and Explainer windows.
6. Capture screenshots, or let the audio translator listen — results appear in the overlay.

---

### 🗝️ Configuration

All settings are stored in:

```
Settings/
 ├── .env                     ← API keys and defaults
 ├── control.ini              ← UI layout and last profile
 ├── glossaries/
 │    ├── default/jp2en.txt   ← Custom term mappings (JP → EN)
 │    └── default/en2en.txt   ← Phrase consistency (EN → EN)
 ├── prompts/                 ← Custom translation/explanation prompts
 └── profiles/                ← Saved overlay color schemes
```

Each overlay remembers its size, position, and colors between sessions.

---

### ⚡ Providers & Models

| Purpose                | Provider                                             | Default Model                                         |
| ---------------------- | ---------------------------------------------------- | ----------------------------------------------------- |
| Screenshot Translation | OpenAI / Gemini                                      | `gpt-4o`, `gemini-3.5-flash`                          |
| Audio Transcription    | OpenAI Whisper / Gemini Audio / Local faster-whisper | `gpt-4o-mini-transcribe`, `gemini-2.5-flash`, `small` |
| Audio Translation      | OpenAI / Gemini                                      | `gpt-4o-mini`                                         |
| Explanation            | OpenAI / Gemini                                      | `gpt-4o-mini`, `gemini-2.5-flash`                     |

---

### 🖼️ Fonts & Visuals

The overlays use **[PixelMPlus](https://itouhiro.github.io/mplus-fonts/)**, licensed under the **Open Font License (OFL 1.1)**.
Place additional `.ttf` fonts in the `fonts/` folder to customize appearance.

The default app icon was sourced from **[Flaticon.com](https://www.flaticon.com/)** and used under their free license.
Please credit the original artist Miguel C Balandrano in derivative works.

---

### 🧠 Local ASR (Optional)

If you select **Local Audio (faster-whisper)** in the Control Panel:

* The script automatically imports `faster-whisper` (CTranslate2)
* Requires the **VC++ 2015–2022 x64 runtime** (auto-installer included)
* Model can be set via `FW_MODEL_NAME` (tiny/base/small/medium/large-v3)
* `FW_COMPUTE` supports `auto`, `int8`, `float16`, or `float32`

GPU (CUDA) use is automatic if available.

---

### 🔍 Glossary Example

`Settings/glossaries/default/jp2en.txt`

```
魔導士 → Mage
神殿 → Temple
癒し → Healing
```

`Settings/glossaries/default/en2en.txt`

```
HP → Health
MP → Mana
```

---

### 🧑‍💻 Contributing

Contributions are welcome!
Developers can improve the translation pipeline, add provider support, or enhance the AutoHotkey UI.

Typical areas for contribution:

* Better VAD (voice detection) logic
* OCR integration improvements
* Caching / offline fallback
* UI performance and cross-window sync

Please open an issue or pull request with clear descriptions and test cases.

---

### ⚖️ Licenses & Credits

| Component                                                                                                                        | License                             | Attribution                                                                                           |
| -------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------- | ----------------------------------------------------------------------------------------------------- |
| JRPG Translator source code                                                                                                      | MIT License                         | © Philipp Reichel                                                                                     |
| [AutoHotkey](https://www.autohotkey.com/)                                                                                        | GNU GPLv2                           |                                                                                                       |
| [Python](https://www.python.org/)                                                                                                | PSF License                         |                                                                                                       |
| [PixelMPlus Font](https://itouhiro.github.io/mplus-fonts/)                                                                       | SIL Open Font License 1.1           |                                                                                                       |
| [Flaticon Icon](https://www.flaticon.com/)                                                                                       | Free License (attribution required) |                                                                                                       |
| [Microsoft Visual C++ Redistributable 2015–2022 (x64)](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist) | © Microsoft Corporation             |                                                                                                       |
| Optional dependencies:                                                                                                           | —                                   | `openai`, `google-generativeai`, `faster-whisper`, `soundcard`, `soundfile`, `numpy`, `python-dotenv` |

---

### 📜 License

This project is licensed under the **MIT License** — see `LICENSE` for details.

```
Copyright (c) 2025 Philipp Reichel

Permission is hereby granted, free of charge, to any person obtaining a copy...
```

---
