# JRPG Translator v0.8.5

v0.8.5 makes overlay placement practical without putting down the controller,
fixes repeat translation requests for unchanged screenshots, and adds a preview
LaunchBox / Big Box integration.

## Highlights

- Translator and Explainer now have dedicated `Move / Resize` modes in their
  appearance tabs.
- XInput controllers work directly in adjustment mode: the left stick moves the
  overlay and the right stick resizes it, with precise low-tilt control and
  faster movement at full tilt.
- Arrow keys provide a controller-mapper fallback. Hold the configured
  Screenshot + Translate key while pressing arrows to resize.
- Enter saves the new window bounds; Escape restores the previous bounds. The
  control panel returns automatically after either action.
- Translation, explanation, overlay, and wheel hotkeys are isolated while an
  overlay is being adjusted, preventing mapped controller inputs from launching
  unrelated actions.
- Repeating Screenshot + Translate with the same capture, visible text, and
  model now completes normally instead of leaving the request glyph active.
- A preview LaunchBox / Big Box plugin is now included under
  `integrations/launchbox`.
- The plugin adds per-game JRPG Translator enablement, JoyToKey profile
  selection, automatic background startup and cleanup, and restoration of the
  previously active JoyToKey profile.
- Plugin setup works from LaunchBox and Big Box, supports portable path storage
  and browse controls, and includes reproducible build, smoke-test, and ZIP
  packaging scripts.

## Install Or Update

1. Download and extract the complete v0.8.5 portable ZIP.
2. Run `JRPG Translator.exe`.
3. Keep a backup of your existing `Settings` folder when updating.

The LaunchBox integration is optional and distributed separately from the main
portable application. Extract its packaged folder into `LaunchBox\Plugins`,
restart LaunchBox, then choose `JRPG Translator Setup...` for a game. Source and
packaging instructions are in `integrations/launchbox/README.md`.

See `CHANGELOG.md` for the complete list of changes.
