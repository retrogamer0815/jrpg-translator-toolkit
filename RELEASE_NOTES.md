# JRPG Translator v0.9.2

Version 0.9.2 is a usability and reliability update for the controller-first
0.9 workflow. It makes the control panel behave naturally at smaller sizes,
finishes the Terminology Overrides on/off workflow, and prevents unrelated
glossary names from appearing as translated speakers.

## Highlights

### Better compact-window behavior

The control panel now has preferred horizontal and vertical resize points. When
resizing close to either point, the corresponding edge settles into the ideal
size without moving the window elsewhere on screen.

The footer keeps its proper height and opaque background while moving over the
scrollable tab content as the window becomes shorter. A vertical scrollbar
appears only after shrinking below the preferred height. Narrow windows wrap
controller guidance instead of clipping it, and squeezed top-level tabs no
longer appear a second time below the tab row.

### Safer terminology overrides

The Terminology Overrides tab now includes **Use terminology overrides**. Turning
it off prevents JP -> TL entries from being sent to translation or explanation
models and prevents local TL -> TL replacement. The setting takes effect
immediately and is stored in Profiles, including Profiles selected through the
LaunchBox / Big Box integration.

JP -> TL instructions now define every entry as a strict conditional exact
match. Screenshot translations also validate translated speaker headers against
the Japanese transcript. If a target belonging to a different Japanese source
is detected, the request is retried once with only glossary entries actually
present in the transcript.

### Controller-only editing

Controller B / Circle can close prompt and terminology text editors, so small
changes no longer require reaching for a mouse just to leave the editor.

Translation-only prompts now format speaker headers consistently, improving
speaker-name color detection.

### Cleaner saving behavior

The manual **Save** button now lives in the optional Paths tab, where it remains
available for manually edited paths and Debug mode. The footer's **Always on
top** setting saves automatically, and redundant dirty-state handling has been
removed.

### Additional fixes

- Reorganized Controller inputs to fit the preferred control-panel height.
- Fixed footer transparency and control-order artifacts while resizing.
- Fixed an artifact covering part of the Audio Translation output-language
  dropdown.

## Included source

The repository contains the AutoHotkey control panel and overlay, Python
translation scripts, bundled multilingual prompts, and LaunchBox / Big Box
plugin source. Personal API keys, Profiles, runtime settings, compiled
executables, backups, logs, screenshots, and generated user data are not
included.

For the complete change list, see `CHANGELOG.md`.
