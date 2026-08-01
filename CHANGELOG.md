# Changelog

All notable changes to JRPG Translator are documented here.

## 0.9.3 - 2026-08-02

This release expands terminology management, improves controller-only setup,
makes borderless overlays easier to close, and adds clearer recovery and error
feedback across screenshot, audio, and explanation workflows.

### Terminology override management

- Replaced raw-text editing as the normal workflow with a two-column terminology
  table showing the detected/model output on the left and its replacement on the
  right.
- Added dedicated Add, Edit, and Delete actions with validation for empty,
  malformed, and duplicate source entries.
- Kept a raw repair editor available when an existing glossary contains lines
  that cannot be represented safely in the table.
- Separated JP -> TL and TL -> TL profile creation and deletion. Each glossary
  type now has its own independent profile list instead of implicitly creating
  or removing the other type.
- Renamed the profile controls to make Manage Entries, New Profile, and Delete
  Profile behavior explicit.
- Reordered the tab to present local TL -> TL correction first and explain that
  these entries can be collected during play or added preemptively.
- Clarified that JP -> TL mappings are sent to the selected model and may be
  ignored or applied unexpectedly depending on the model and prompt complexity.
- Made terminology warnings and validation messages owned dialogs so they remain
  above the manager and entry windows.

### Screenshot translation safeguards

- Added local detection for malformed mixed-script output such as partially
  translated names containing both Latin and Japanese characters.
- Detects a JP -> TL target used in dialogue when its exact Japanese source is
  absent from the transcript.
- Suspicious glossary-influenced output is retried once with only exact glossary
  sources found in the transcript; the original usable response remains the
  fallback if the corrective request fails.
- Distinguished a missing capture target from a PNG size limit that is too low.
  The overlay now gives a dedicated explanation and recommends increasing the
  configured limit instead of incorrectly reporting that no target was set.
- Improved progressive PNG reduction and reports the smallest attempted image
  when the configured limit still cannot be reached safely.

### Overlay access and request feedback

- The control-panel title-bar X now performs the same complete shutdown as
  Close all, preventing hidden overlays from being left running accidentally.
- Added discreet borderless `...` and `×` controls that appear when the pointer
  enters either overlay, without adding a title bar or disturbing overlay text.
- Renamed the context-menu exit action to identify the Translator or Explainer
  overlay it closes.
- Added a `Generating explanation...` notification as soon as an explanation
  request is accepted, followed by the existing completion or failure feedback.
- Missing OpenAI or Gemini keys now produce a provider-specific message in the
  appropriate overlay for screenshot translation, live audio, and explanations.
  Requests are stopped before a worker starts, and Python-level fallbacks cover
  independently launched scripts.

### Live audio reliability

- Added automatic reconnection with bounded backoff for temporary DNS, network,
  WebSocket, rate-limit, and service-availability failures.
- Authentication and other configuration failures remain immediate errors
  instead of being retried indefinitely.
- Replaced per-second WMI/COM process scans with direct PID tracking. WMI is now
  used only as a one-time recovery fallback, preventing the AutoHotkey memory
  failure seen after leaving the application unattended.

### API setup and application information

- Reorganized API Keys into a recommended Windows Environment Variables section
  and a clearly separated optional in-app `.env` section.
- Added a button that opens the Windows Environment Variables editor directly
  and updated the accompanying setup instructions.
- Added an About dialog with the application version, author, license, GitHub
  and bug-report links, X contact, and copyable diagnostic version information.

### Controller and interface refinements

- Keyboard/controller arrows and the controller D-pad can now switch between
  Keyboard inputs and Controller inputs in the Controls tab.
- Pressing Down on Keyboard inputs now enters the first shortcut field instead
  of skipping the shortcut list and jumping to the footer.
- Standardized both overlay sliders on the term `Opacity` and expanded `Max PNG
  size` to `Maximum PNG size` for clearer alignment.
- Moved About into the API action row so it remains visible at the preferred
  magnetic-snap width and height.

### LaunchBox / Big Box integration and documentation

- Added `Open JRPG Translator...` to the per-game setup window for first-time
  API, hotkey, capture, and overlay configuration, and for bringing an existing
  control panel forward.
- Added concise first-time guidance beside the button and enlarged/rearranged
  the setup window so all explanatory text remains visible.
- Added a visual README showcase for screenshot translation, transcripts with
  kanji readings, explanations, live audio translation, and LaunchBox setup,
  including an updated plugin screenshot.

## 0.9.2 - 2026-07-26

This maintenance release improves compact-window behavior, finishes the
Terminology Overrides workflow, strengthens controller-only editing, and
prevents unrelated glossary names from leaking into screenshot translations.

### Compact and resizable control panel

- Reorganized Controller inputs so its direct-binding and D-pad options fit
  beside the binding controls without making the tab taller than Keyboard
  inputs at the preferred window size.
- Made the controller guidance wrap vertically as the window narrows instead of
  being clipped off the right edge.
- Fixed squeezed top-level tabs being rendered a second time below the tab row.
- Added magnetic horizontal and vertical resize points at the preferred control
  panel size without moving the window on screen.
- Kept the footer at a fixed height while allowing it to move upward over
  scrollable tab content as the window shrinks.
- Delayed the vertical scrollbar until the window is reduced below its preferred
  height.
- Gave the footer an opaque, correctly ordered background so content beneath it
  cannot show through.
- Fixed a stray rectangle covering part of the Audio Translation output-language
  dropdown.

### Terminology Overrides

- Added a persistent Use terminology overrides option to enable or disable both
  JP -> TL and TL -> TL processing throughout supported workflows.
- Stored the option in Profiles so Profiles selected through LaunchBox / Big Box
  apply the same terminology policy.
- Made changes effective immediately for open Translator and Explainer
  workflows.
- Strengthened JP -> TL model instructions so entries are strict conditional
  exact matches and absent Japanese sources cannot introduce glossary targets.
- Added local speaker-header validation. A glossary target assigned to a
  different Japanese speaker triggers one retry using only entries found in the
  transcript.
- Exact Japanese speaker matches now force their configured target locally;
  failed corrective retries fall back safely to kana romanization or the
  original Japanese name.

### Controller and editing improvements

- Controller B / Circle can now close prompt and terminology text editors after
  navigating to their Close button.
- Standardized speaker headers in translation-only prompts so speaker-name color
  detection works consistently across prompt families.

### Saving and state cleanup

- Moved the manual Save action into the optional Paths tab, where it remains
  available for edited paths and Debug mode.
- Made the footer's Always on top option save automatically.
- Removed obsolete dirty-state wiring left behind by settings that already save
  immediately.

## 0.9.1 - 2026-07-24

This maintenance release improves in-game usability, makes audio capture easier
to diagnose, and resolves several controller, prompt, glossary, and dialog
issues found after the 0.9.0 release.

### Audio capture diagnostics

- Added a Test Audio button to the Audio Translation tab.
- The test checks the selected Windows output locally without making an API
  request and reports whether audible output is reaching the capture device.
- Added concise troubleshooting guidance for device selection and emulator
  audio drivers such as XAudio and WASAPI.
- Fixed opening the Windows default output and explicitly selected endpoints
  during the audio test.

### Prompt and output handling

- Removed the normal-user Translation post-processing selector.
- Translation prompts containing `with_transcript` or `with_kanji_reading`
  now select transcript-aware output handling automatically.
- Added a hidden Direct model output override to the optional Paths tab; it is
  disabled by default.
- Cleaned up translation-only prompts and synchronized multilingual explanation
  prompts.
- Explanation prompts now request hiragana readings for kanji in the repeated
  Japanese context.
- Restored TL -> TL terminology replacement for screenshot translations while
  keeping the same case-insensitive matching and case-preserving replacement
  behavior used by explanations and live audio output.

### Controller and keyboard behavior

- Added a persistent Use D-pad for control panel navigation option.
- Stored the D-pad navigation preference in Profiles so LaunchBox / Big Box can
  apply it with the rest of a setup.
- When D-pad navigation is disabled, controller A / Cross and B / Circle remain
  available while mapper-generated keyboard arrows can navigate without
  duplicate tab movement.
- Changed Esc to behave like controller B / Circle for cancel and back actions
  instead of hiding the entire control panel.
- Removed automatic JoyToKey-process inference from control-panel navigation.
- Improved controller behavior for sliders, numeric fields, capture dialogs,
  model pickers, color controls, and overlay move / resize modes.

### Interface and dialog polish

- Added persistent opacity control for the complete control panel, including its
  controls and title bar.
- Added dark-mode support to prompt-name, Profile-name, and hotkey-entry
  dialogs.
- The New prompt dialog now documents the `with_transcript` and
  `with_kanji_reading` naming convention.
- Fixed Controls-tab label and description repaint artifacts.
- Refined Screenshot Translation spacing after removing the post-processing
  selector.

### Reliability fixes

- Fixed destroyed-control errors when closing or refreshing online model lists.
- Fixed warnings and duplicate navigation caused by mixed controller and
  keyboard input paths.
- Fixed stale or clipped capture and move / resize instruction overlays.
- Fixed immediate controller cancellation in capture and overlay-adjust modes.
- Fixed repeat screenshot requests, output color restoration, caret hiding, and
  several overlay refresh edge cases.

## 0.9.0 - 2026-07-23

This release unifies per-game configuration into Profiles, expands native
controller support across the control panel and capture tools, and turns the
LaunchBox / Big Box integration into a practical controller-first launcher.

### Unified Profiles

- Added a dedicated Profiles tab for saving and applying complete setups.
- Profiles store the selected screenshot and explanation prompts, translation
  post-processing, terminology profiles, capture region or window, and both
  overlays' size, position, colors, transparency, font, size, and weight.
- Replaced the separate Translator and Explainer appearance-profile controls
  with the unified profile workflow.
- Applying a Profile updates open overlays safely, keeps their saved positions
  on screen, and supports background startup through `--profile`.
- Added per-game Profile selection to the LaunchBox / Big Box plugin.

### Native controller controls

- Renamed the Hotkeys tab to Controls and added separate Keyboard inputs and
  Controller inputs views.
- Added optional direct XInput action bindings without removing keyboard
  hotkeys or compatibility with JoyToKey and similar mapping tools.
- Added native D-pad navigation plus controller A / Cross confirmation and
  B / Circle cancellation throughout the control panel, including duplicate
  input suppression when a controller mapper also emits arrow keys.
- Added controller operation for model dialogs, font size, Max PNG size,
  transparency, color sliders, dropdowns, and other compact editors.
- Added accelerated held-input behavior for color and transparency sliders.

### Controller capture setup

- Added a controller-first Capture > Region mode. The left stick moves the
  current region and the right stick resizes it; A saves and B cancels.
- Added a controller-first Capture > Window mode with directional window
  cycling, selected-window highlighting, and safe foreground previews.
- Preserved the conventional drag-to-select Region flow when capture setup is
  opened with a mouse.
- Improved capture-mode dialogs, dark-mode styling, spatial navigation,
  instruction HUD layout, cancellation, click-through cleanup, and right-stick
  axis consistency.

### Overlay and appearance refinements

- Extended direct-controller support for moving and resizing Translator and
  Explainer overlays, including immediate A / B confirmation and cancellation.
- Added controller-friendly font-size adjustment and clearer active-control
  indicators.
- Improved the hue, saturation, and brightness editor with informative
  gradients, live previews, dark-mode focus frames, and faster held movement.
- Fixed recurring caret, focus, color, and overlay-HUD issues during resizing,
  scrolling, profile changes, and appearance updates.

### LaunchBox / Big Box integration

- Added per-game JRPG Translator Profile selection alongside JoyToKey profiles.
- Fixed cold startup so configured games launch JRPG Translator even when it is
  not already running.
- Made JRPG Translator and JoyToKey independently optional per game.
- Added direct-controller navigation in Big Box while retaining normal mouse
  dialogs in LaunchBox.
- Added a Big Box-native controller file and folder browser for locating the
  Translator executable, JoyToKey executable, and JoyToKey profiles.
- Improved directional navigation, dropdown reliability, focus styling,
  theme contrast, window sizing, explanatory text, and the context-menu icon.

### Prompts and model workflow

- Updated every screenshot-translation prompt to combine multiple screenshots
  in capture order before translating the complete reconstructed passage.
- Rebuilt all localized explanation prompts from the updated English learning
  prompt while preserving language-specific instructions.
- Improved controller focus and cleanup in online model discovery dialogs,
  including correct action-button highlighting and safe cancellation.

### Reliability

- Fixed destroyed-control warnings after closing model, color, capture, and
  adjustment dialogs.
- Fixed capture selection leaving an invisible input-blocking overlay behind.
- Improved native-controller polling and transient input handling across modal
  dialogs and overlay adjustment modes.

## 0.8.5 - 2026-07-21

This release adds controller-driven overlay positioning, fixes repeat requests
for unchanged screenshots, and introduces a preview LaunchBox / Big Box
integration with per-game JoyToKey support.

### Controller Move / Resize

- Added `Move / Resize` commands to the Translation Window and Explanation
  Window tabs.
- Added direct XInput support: the left analog stick moves the selected overlay
  and the right analog stick resizes it, with analog speed scaling for precise
  small adjustments and faster full-tilt movement.
- Added keyboard fallback controls. Arrow keys move the overlay; holding the
  configured Screenshot + Translate key while pressing arrows resizes it.
- Enter saves the new bounds, Escape restores the previous bounds, and the
  control panel returns automatically after either action.
- Temporarily suppresses translation, explanation, overlay, and wheel hotkeys
  while adjustment mode is active so controller mappings cannot trigger an
  unrelated workflow.
- Added on-screen mode guidance and success/cancel notifications while keeping
  the game from being activated unnecessarily.

### Screenshot request reliability

- Added a separate completion marker for screenshot OCR output so requesting
  the same capture, text, and model twice is recognized as a new completed
  request instead of leaving the busy glyph visible indefinitely.

### LaunchBox / Big Box integration preview

- Added the LaunchBox plugin source under `integrations/launchbox`.
- Added `JRPG Translator Setup...` to the per-game LaunchBox context menu and
  Big Box game details menu.
- Added per-game controls for enabling JRPG Translator and selecting a JoyToKey
  profile.
- Starts JRPG Translator in background mode before configured games and closes
  only instances started by the plugin after gameplay.
- Starts JoyToKey or switches an existing instance to the selected profile,
  then restores the previous profile when the game exits.
- Added browse controls for the Translator executable, JoyToKey executable, and
  JoyToKey profile folder, with portable relative path storage where possible.
- Improved Big Box profile-selector contrast and keyboard/controller operation.
- Added repeatable build, smoke-test, and installable ZIP packaging scripts.

## 0.8.0 - 2026-07-20

This release completes the controller-first configuration workflow, adds
provider-backed model discovery, and expands terminology profiles across every
translation path.

### Model management

- Separated model lists for Screenshot Translation and Explanation so each
  workflow can keep models suited to its own input type and purpose.
- Added an `Add model` flow that can browse models available to the configured
  OpenAI or Gemini API key or accept a model ID manually.
- Added task-aware online filtering for screenshot-capable, explanation, and
  live-audio models, with natural sorting and a local catalog cache.
- Added controller-friendly model dialogs with arrow navigation, Enter-based
  selection, dark-mode styling, refresh controls, and direct movement from the
  model list to its action buttons.
- Disabled model controls for the provider that is not currently selected.

### Controller and appearance controls

- Added controller operation for overlay transparency and color settings.
- Added a custom hue, saturation, and brightness editor with live color preview,
  informative gradient tracks, and clear light/dark focus indicators.
- Added independent Bold controls for Translator and Explainer overlay fonts.
- Improved RichEdit font handling for Japanese-capable pixel fonts, including
  PixelMplus10 and PixelMplus12, while retaining speaker-name colors.
- Added dark-mode support to glossary and prompt text editors and stopped their
  contents from opening fully selected.
- Hid the advanced Paths tab by default behind a `showPathsTab` setting and made
  Debug mode default to off while preserving both options for advanced users.

### Terminology overrides

- Changed replacement matching to be case-insensitive while preserving the
  exact replacement text entered by the user.
- Applied JP-to-target-language and target-language-to-target-language glossary
  rules to screenshot translations and explanations.
- Applied target-language-to-target-language glossary rules to live-audio
  translation output.

### Interface and reliability

- Refined alignment, spacing, dropdown widths, and button sizing across the
  Screenshot Translation, Translation Window, Explanation, and Explanation
  Window tabs.
- Unified the Translation Window and Explanation Window appearance layouts and
  kept compact color swatches aligned with the other controls.
- Fixed Debug mode propagation to the Python processes and overlays.
- Fixed additional caret, text-color, font-refresh, and control-lifetime issues
  found while changing overlay appearance or closing controller dialogs.

## 0.7.5 - 2026-07-19

This release focuses on emulator-friendly startup, a more compact live-audio
workflow, responsive control-panel sizing, capture reliability, and overlay
polish.

### Control panel and audio

- Replaced the separate Start Audio, Toggle Listening, and status area with one
  compact `Audio Translation On/Off` button that reflects the active state.
- Added non-activating popup confirmations when audio translation starts or
  stops, including when triggered from a controller-mapped hotkey.
- Added a scrollable 890 x 680 design canvas so the control panel can be resized
  down to approximately 640 x 480 without losing access to controls.
- Added native horizontal and vertical scrollbars, mouse-wheel scrolling, and
  dark-mode scrollbar styling.
- Made the custom tab row fit narrower windows while keeping every tab
  reachable and preserving the beginning of shortened labels.
- Fixed ghost tab controls, alternating page offsets, stale scroll positions,
  and clipped tab labels in compact layouts.
- Standardized dropdown widths in Screenshot Translation and Audio Translation
  and reduced unused footer space.
- Fixed initially selected-looking Font and Profile dropdown text and related
  invalid-value startup warnings.
- Added natural model sorting so model families and version numbers remain
  grouped in a predictable order.
- Added built-in OpenAI model entries for `gpt-5.5`, `gpt-5.4-nano`, and
  `gpt-5.4-pro`.
- Added built-in Gemini model entries for `gemini-3.1-flash-lite`,
  `gemini-3.5-flash`, and `gemini-3.1-pro-preview`.

### Startup and emulator integration

- Added a background-start mode for front ends such as Big Box so launching the
  translator alongside a game does not expose the control panel or foreground
  the Windows taskbar.
- Improved startup focus handling so emulators retain control while configured
  overlays open automatically.
- Allowed Translator to start visible and always on top while Explainer starts
  independently in the background when it is not configured as always on top.
- Simplified Translator and Explainer appearance settings to one window color
  while retaining a Windows 10 border fallback.

### Overlay behavior

- Prevented overlays from taking focus or restoring a blinking text caret when
  opened through the control panel or scrolled while inactive.
- Improved controller and mouse-wheel scrolling of inactive overlays without
  swallowing wheel input in unrelated Windows applications.
- Preserved Translator text and speaker-name colors when opening Explainer or
  applying live font, color, transparency, and window-theme changes.
- Improved request-status glyph size and placement so it does not obscure text
  near the upper-right corner.
- Cleaned up explanation-ready and audio-state notifications.
- Corrected independent Translator and Explainer startup visibility, z-order,
  and always-on-top behavior.

### Capture and translation

- Capture commands now start Translator automatically when the overlay is not
  already running and wait for it to become ready before selection begins.
- Improved capture-window highlighting after moving between the desktop, full
  screen, and individual application windows.
- Added the least restrictive Gemini safety settings supported by the API for
  mature game translation and explanation requests.
- Retained OpenAI temperature compatibility for models that reject custom
  temperature values while using temperature 0 where supported.

### Terminology, naming, and documentation

- Reworded Terminology Overrides around language-neutral `JP -> TL` and
  `TL -> TL` glossary profiles, where TL means the prompt-selected target
  language.
- Expanded the terminology introduction and examples for multilingual use.
- Standardized the source and executable names as `JRPG Translator.ahk`,
  `JRPG Translator.exe`, `bin/overlay.ahk`, and `bin/overlay.exe`.
- Updated source-build paths, component names, workflow instructions, and
  release documentation.

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
