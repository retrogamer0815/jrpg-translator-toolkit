# JRPG Translator v0.9.4

Version 0.9.4 introduces the **Study Library** and **Study Reader**, turning
saved explanations into an organized, searchable collection with their source
screenshots and gameplay context. The release also improves controller-first
operation, low-resolution compatibility, LaunchBox setup, and interface
polish.

## Highlights

### A persistent Study Library

The Explanation tab can now save explanations to a local SQLite Study Library.
Each entry keeps the original Japanese, complete explanation, parsed sections,
Profile, provider, model, prompt, metadata, and optional source screenshots.

Repeated explanations of the same Japanese source are grouped as versions, so
trying a second or third explanation no longer creates unrelated entries.
Optional plain-text copies remain available and use `_v02`, `_v03`, and later
suffixes for repeated attempts.

The Library provides:

- Search and filters for Profile, chapter, speaker, tags, date/time, and Anki
  status.
- Sortable, configurable columns with remembered widths.
- Bulk editing of chapter, speaker, tags, and `Added to Anki` state.
- Optional source screenshots and storage-size information.
- Multiple named libraries that can be created, switched, renamed, archived,
  and restored.
- Profile-based library selection, including Profiles loaded by the LaunchBox
  plugin.

### A dedicated Study Reader

The new Reader puts the explanation itself first and keeps screenshots,
original Japanese, and metadata available as context. It supports navigation
between entries, versions, screenshots, and parsed sections such as the natural
translation, detailed analysis, vocabulary, nuance, and takeaways.

Explanations can be edited by section or as a complete document. Manual edits
are marked and can be reverted to the original model response.

A new **Copy...** menu copies the current section or complete explanation. When
viewing Original Japanese, it can also copy the text without attached hiragana
readings—for example, `お城(しろ)には行(い)けました？` becomes
`お城には行けました？`—while leaving ordinary parentheses intact.

Standalone Study Library and Study Reader launchers are included for reviewing
material without opening the main control panel first.

### Better explanation saving and provider selection

The Explanation tab now separates:

- **Save explanations to Study Library**
- **Include source screenshots in Study Library**
- **Save plain-text copies**

Plain-text copies are sorted under the active unified Profile. Without an active
Profile, they remain in the main `Settings\Explanations` folder.

An intermittent provider-state bug could incorrectly request an OpenAI key even
when Gemini was selected. Requests now synchronize the visible provider and
model before credential validation.

### Controller-first control-panel navigation

- LB / L1 and RB / R1 move between the main tabs while the control panel is
  active.
- Translation/explanation actions assigned to those buttons are suppressed
  while the panel is being navigated, including matching JoyToKey hotkeys.
- Assignment dialogs suppress tab switching while capturing a controller
  button.
- Held arrow keys sent by JoyToKey now repeat naturally in sliders and color
  controls.
- Mapped numpad keys can no longer overwrite **Maximum PNG size** merely while
  navigating past the field.

Temporary screenshots selected for deletion are now cleaned at application
startup, after the previous session had a chance to import them into the Study
Library.

### Better 720p and high-scaling support

The main control panel, Study Library, Study Reader, and LaunchBox setup window
have received additional high-DPI and low-resolution work. The LaunchBox window
uses a scrollbar only when needed, keeps Save and Cancel reachable, and no
longer skips the JoyToKey section while scrolling.

Profile dropdowns in the plugin now correctly commit keyboard/controller
selections with Enter or Space.

### Interface and reliability polish

- Improved dark-mode coverage across Study windows and dialogs.
- Remembered Study Library and Reader size, position, columns, and reading
  state.
- Reduced startup flashes, window jumping, resize trails, and stale screenshot
  placement.
- Fixed the first Study Library double-click after opening.
- Hardened redraw timers and shutdown callbacks against closed or destroyed
  controls.
- Prevented the brief themed-dialog warning that could appear while exiting.

## Compatibility and upgrade notes

- Existing v0.9.3 settings and unified Profiles remain usable.
- The default database is created locally at
  `Settings\Study Library\study_library.db` when Study Library saving is used.
- Additional libraries are kept under `Settings\Study Libraries`.
- No existing plain-text explanation files are removed.
- API requests still use the selected OpenAI or Gemini provider; the Study
  Library itself is local and does not require another service.

## Included source

The repository includes the AutoHotkey control panel and overlays, Python
translation/explanation and Study Library helpers, standalone Study launchers,
and the LaunchBox / Big Box integration source.
