# JRPG Translator v0.9.0

v0.9.0 brings the controller-first workflow together. Complete setups can now
be saved as Profiles, capture targets can be configured without a mouse, and
the LaunchBox / Big Box plugin can apply the right Translator and JoyToKey
profiles automatically for each game.

## Highlights

- New unified Profiles store prompts, translation post-processing, terminology
  overrides, capture target, and both overlays' complete appearance and bounds.
- Profiles can be applied while the overlays are running and selected per game
  from the LaunchBox / Big Box plugin.
- The Controls tab now supports both keyboard hotkeys and optional direct
  controller action bindings.
- Native D-pad navigation plus A / Cross and B / Circle operation works
  throughout the control panel without requiring JoyToKey.
- Capture > Region can be moved and resized with analog sticks, while
  Capture > Window can cycle and preview available windows with a controller.
- Controller operation now covers overlay move/resize, font size, Max PNG size,
  transparency, color editing, model discovery, and compact modal dialogs.
- The LaunchBox / Big Box plugin now supports Translator Profiles, reliable
  cold startup, independent JoyToKey use, improved navigation and styling, and
  a controller-native Big Box path browser.
- Screenshot prompts combine multiple captures in order before translating the
  reconstructed passage.
- All localized explanation prompts have been rebuilt from the latest English
  Japanese-learning prompt.
- Numerous focus, caret, capture-overlay, modal cleanup, and duplicate-input
  issues have been fixed.

## Install Or Update

1. Download and extract the complete v0.9.0 portable ZIP.
2. Keep a backup of your existing `Settings` folder when updating.
3. Run `JRPG Translator.exe`.

The optional LaunchBox integration is distributed separately from the main
portable application. Extract its packaged folder into `LaunchBox\Plugins`,
restart LaunchBox, then choose `JRPG Translator Setup...` for a game. Source and
packaging instructions are in `integrations/launchbox/README.md`.

See `CHANGELOG.md` for the complete list of changes.
