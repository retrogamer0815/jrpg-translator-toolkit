# Changelog

All notable changes to JRPG Translator are documented here.

## 0.7.5 - In development

This is the current development source for the next JRPG Translator release.

### Control panel and audio

- Replaced the separate Start Audio, Toggle Listening, and status area with one
  compact `Audio Translation On/Off` button that reflects the active state.
- Added non-activating popup confirmations when audio translation starts or
  stops, including when triggered from a controller-mapped hotkey.
- Added a background-start mode for front ends such as Big Box so launching the
  translator alongside a game does not expose the control panel or foreground
  the Windows taskbar.
- Improved startup focus handling so emulators retain control while configured
  overlays open automatically.
- Simplified Translator and Explainer appearance settings to one window color
  while retaining a Windows 10 border fallback.
- Reworded Terminology Overrides around language-neutral `JP -> TL` and
  `TL -> TL` glossary profiles, where TL means the selected target language.

### Overlay behavior

- Standardized the compiled overlay name as `bin/overlay.exe` while source runs
  continue to use the AutoHotkey overlay script.
- Prevented the Translator overlay from taking focus or restoring a blinking
  text caret when opened through the control panel.
- Corrected independent Translator and Explainer startup visibility and
  always-on-top behavior when both share the same overlay implementation.
- Preserved existing Translator colors and speaker-name formatting when the
  Explainer is opened or recreated.
- Improved request-status glyph placement so it does not obscure translated
  text near the upper-right corner.
- Cleaned up explanation-ready notifications and improved overlay theme refresh
  isolation.

### Documentation

- Updated source-build paths and component names for the standardized overlay.
- Clarified multilingual terminology overrides and their profile format.

## 0.7.0 - 2026-07-13

This is a major controller-usability, live-audio, overlay, and interface update
over v0.6.1.

### Live audio translation

- Replaced the old transcription-then-translation pipeline with direct live
  translation through OpenAI Realtime Translate and Gemini Live Translate.
- Removed the local faster-whisper workflow, VAD tuning controls, VC++ runtime
  installation check, and legacy transcription interface.
- Added output-language selection for live audio translation.
- Kept Windows playback-device selection and refresh controls in a simpler
  Audio Translation tab.
- Added automatic bottom-following for new live-audio subtitles while retaining
  scrollback for earlier lines.

### Controller-first control panel

- Added spatial arrow-key navigation based on the visible control layout.
- Added Enter activation for buttons, checkboxes, and dropdown selection.
- Added controller-friendly dropdown behavior: open, move with Up/Down, confirm
  with Enter, and cancel with Escape.
- Added Page Up/Page Down tab switching for mapping to controller shoulder
  buttons.
- Replaced the native tab header with a custom tab bar that clearly distinguishes
  the active tab and controller navigation focus.
- Added stronger focus styling for the selected control.
- Corrected navigation order across model rows, action rows, tab pages, and the
  persistent bottom action bar.
- Prevented ordinary mouse clicks on dropdowns and buttons from being mistaken
  for tab clicks.
- Fixed dropdowns that previously needed a second mouse click after focus
  styling was applied.

### Overlay behavior

- Removed the blinking RichEdit caret while preserving wheel and controller
  scrolling.
- Allowed visible Translator and Explainer overlays to react to mapped scrolling
  input without taking focus from the game or emulator.
- Added non-activating show behavior so overlays do not trigger emulator
  "pause when inactive" behavior.
- Reset screenshot translations and explanations to the top when new content
  arrives.
- Kept live-audio output pinned to the newest visible subtitle at the bottom.
- Added per-overlay busy glyphs for screenshot translation and explanation
  requests without clearing the previous content.
- Ensured screenshot requests show their glyph in Translator and explanation
  requests show it in Explainer.
- Replaced the distracting screenshot confirmation tooltip with the overlay
  glyph.
- Prevented opening Explainer from resetting existing Translator text colors.
- Corrected theme refresh targeting so each newly opened overlay initializes
  independently.
- Added scrolling that works regardless of the mouse-pointer position.

### Control panel and appearance

- Added a persistent soft-charcoal dark mode covering the title bar, tabs,
  buttons, checkboxes, dropdowns, fields, disabled states, and navigation focus.
- Preserved actual overlay color previews while dark mode is active.
- Added separate, consistent Live Input and Live Translation headings.
- Fixed corrupted UTF-8 text in buttons, help text, hotkey controls, and
  notifications.
- Added a cleaner custom tab bar with reliable mouse and controller behavior.
- Improved selected-control visibility throughout the interface.

### Screenshot translation and explanations

- Added maintained default English screenshot prompts with and without hiragana
  readings, plus a default English Japanese-learning explanation prompt.
- Synchronized OpenAI and Gemini model lists between Screenshot Translation and
  Explanation without changing the other tab's active selection.
- Fixed model-list deletion, stale selections, startup warnings, and invalid
  saved-model values.
- Added OpenAI temperature compatibility for models that only accept their
  default temperature while retaining temperature 0 where supported.
- Updated Gemini explanation support to the current `google-genai` package.
- Added configurable speaker-name coloring and optional guessed-subject italics.
- Added transcript and translation post-processing modes.
- Removed the broad "short first Japanese line" speaker-name fallback that
  produced false positives; speaker labels now require explicit brackets.
- Improved prompt profile, glossary, and model synchronization behavior.

### Capture and hotkeys

- Fixed capture-region and capture-window workflows that could bring the control
  panel forward a second time after selection.
- Preserved capture coordinates when unrelated settings are saved.
- Made Screenshot + Translate honor the configured hotkey instead of assuming a
  fixed key combination.
- Improved function-key handling for screenshot glyph notifications without
  adding translation delay.
- Added reliable show/hide behavior for Translator, Explainer, and the control
  panel.

### Reliability

- Added a current source-install dependency manifest in `requirements.txt`.
- Added safer first-run defaults for compiled and source builds.
- Compiled builds now launch the compiled overlay automatically; source builds
  continue to launch the AHK overlay source.
- Fixed repeated AutoHotkey `#Warn` startup dialogs and invalid control values.
- Fixed dark-mode swatches, tab redraws, disappearing controls, and stale focus
  indicators.
- Improved overlay theme, scroll, model, and window-state isolation.

## 0.6.1 - 2025-12-13

- Preserved capture-region coordinates when saving control-panel settings.
- Improved capture completion detection when selecting the same region again.
