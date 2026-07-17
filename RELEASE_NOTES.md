# JRPG Translator v0.7.5 (in development)

This is the current source snapshot for the next JRPG Translator release. More
small refinements may be added before the final v0.7.5 package is published.

## Highlights

- A single `Audio Translation On/Off` button now controls and displays live-audio
  state without a separate listening toggle or status area.
- Audio start and stop actions provide non-activating on-screen confirmations.
- Background startup and overlay launch behavior better preserve emulator focus
  and avoid surfacing the Windows taskbar when launched through a game front end.
- Translator and Explainer overlays now initialize independently while sharing
  the same overlay implementation.
- Overlay window-color settings are simpler, with compatibility handling kept
  for Windows 10 borders.
- Terminology Overrides now use language-neutral `JP -> TL` and `TL -> TL`
  wording for any prompt-selected target language.
- The compiled overlay uses the clearer `bin/overlay.exe` name.

## Development Status

The source files on `main` are labeled for v0.7.5 development. A final portable
ZIP and GitHub release will follow after the remaining refinements are complete.

## Install Or Update

1. Download and extract the complete portable ZIP when v0.7.5 is released.
2. Run `JRPG Translator.exe`.
3. Add API keys in the control panel or through Windows environment variables.

When updating an existing installation, keep a backup of your `Settings`
folder. The release contains maintained `default` prompt profiles, so preserve
any customized files that use the same names. The bundled `example` glossary
does not replace customized glossary profiles.

See `CHANGELOG.md` for the complete list of changes.
