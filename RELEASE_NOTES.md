# JRPG Translator v0.7.0

This is a major usability and audio-translation update.

## Highlights

- Direct live audio translation with OpenAI Realtime Translate or Gemini Live Translate models.
- Controller-first control panel navigation with predictable directional movement, custom tabs, Enter activation, and clearer focus styling.
- Translator and Explainer overlays can remain non-activating so games and emulators keep focus, while controller-mapped scrolling still works.
- Optional dark mode covering the control panel, tabs, controls, labels, and color previews.
- Faster, more reliable screenshot translation and explanation workflows, with synchronized model lists and improved prompt handling.
- Numerous overlay, capture, hotkey, encoding, redraw, and first-run fixes.

## Audio Change

The previous two-stage audio pipeline and local transcription option have been removed. Audio translation now uses a provider's dedicated live translation model directly. This reduces delay and keeps the Audio Translation tab simpler.

## Install Or Update

1. Download and extract the complete portable ZIP.
2. Run `JRPG Translator.exe`.
3. Add API keys in the control panel or through Windows environment variables.

When updating an existing installation, keep a backup of your `Settings`
folder. The release contains maintained `default` prompt profiles, so preserve
any customized files that use the same names. The bundled `example` glossary
does not replace customized glossary profiles.

See `CHANGELOG.md` for the complete list of changes.

