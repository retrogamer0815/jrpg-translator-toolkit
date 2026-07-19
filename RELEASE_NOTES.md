# JRPG Translator v0.7.5

This release focuses on emulator-friendly startup, a more compact live-audio
workflow, responsive control-panel sizing, capture reliability, and overlay
polish.

## Highlights

- A single `Audio Translation On/Off` button now controls and displays live-audio
  state without a separate listening toggle or status area.
- Audio start and stop actions provide non-activating on-screen confirmations.
- The control panel can now be resized below its full design size; native
  scrollbars keep every control and tab accessible in compact windows.
- The tab row adapts to narrow windows and retains readable label beginnings.
- Background startup and overlay launch behavior better preserve emulator focus
  and avoid surfacing the Windows taskbar when launched through a game front end.
- Translator and Explainer overlays now initialize independently while sharing
  the same overlay implementation.
- Inactive overlay scrolling no longer consumes the mouse wheel in unrelated
  Windows applications, and scrolling no longer restores the text caret.
- Capture commands automatically start Translator when necessary, and window
  highlighting follows the currently hovered desktop or application window.
- Overlay window-color settings are simpler, with compatibility handling kept
  for Windows 10 borders.
- Terminology Overrides now use language-neutral `JP -> TL` and `TL -> TL`
  wording for any prompt-selected target language.
- Model lists use natural sorting and include additional current OpenAI and
  Gemini defaults.
- The canonical program names are `JRPG Translator.exe`,
  `JRPG Translator.ahk`, `bin/overlay.exe`, and `bin/overlay.ahk`.

## Install Or Update

1. Download and extract the complete v0.7.5 portable ZIP.
2. Run `JRPG Translator.exe`.
3. Add API keys in the control panel or through Windows environment variables.

When updating an existing installation, keep a backup of your `Settings`
folder. The release contains maintained `default` prompt profiles, so preserve
any customized files that use the same names. The bundled `example` glossary
does not replace customized glossary profiles.

See `CHANGELOG.md` for the complete list of changes.
