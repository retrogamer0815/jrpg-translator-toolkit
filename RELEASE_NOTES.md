# JRPG Translator v0.9.3

Version 0.9.3 expands terminology management and improves the controller-first
workflow while adding clearer feedback and stronger recovery behavior across
screenshot, live-audio, and explanation requests.

## Highlights

### Terminology tables and independent profiles

Terminology overrides no longer need to be maintained as raw
`source -> replacement` text during normal use. **Manage Entries...** opens a
two-column table with Add, Edit, and Delete actions and validation for malformed
or duplicate entries. A raw repair editor remains available for files that
cannot be represented safely in the table.

JP -> TL and TL -> TL profiles are now managed independently. The revised tab
puts local TL -> TL correction first and explains that it is useful for fixing
model output over time. JP -> TL guidance now makes clear that those mappings
are model instructions and can be affected by the model and prompt complexity.

### Safer screenshot output

Screenshot translation now detects strong signs of glossary contamination,
including unrelated glossary targets and mixed Latin/Japanese words such as a
partially translated name. It can retry once using only exact glossary sources
found in the transcript while retaining the original usable result as a safe
fallback.

A capture that cannot be reduced below **Maximum PNG size** now receives its
own explanation instead of the misleading `No target set` message. The overlay
also reports the smallest attempted PNG and recommends a practical limit.

### Easier-to-close overlays

The control-panel title-bar X now closes the complete application, matching
**Close all**. Translator and Explainer overlays gain discreet borderless
`...` and `×` controls when the pointer enters the window, and their context
menus name the specific overlay being closed.

### Clearer request and API-key feedback

Explanation requests show **Generating explanation...** immediately, followed
by the existing completion or failure message.

When the selected provider has no API key, screenshot translation, audio
translation, and explanations now stop before launching their worker and show a
clear OpenAI- or Gemini-specific message in the appropriate overlay. Worker-level
fallbacks provide the same result when a Python script is launched separately.

The API Keys tab now recommends Windows Environment Variables first, opens the
Windows editor directly, and separates optional in-app `.env` storage into its
own section. A new About dialog provides version, author, license, project and
bug-report links, contact information, and copyable diagnostics.

### Live-audio resilience

Temporary DNS, network, WebSocket, rate-limit, and service-availability failures
now reconnect automatically with bounded backoff. Authentication and other
configuration errors still fail immediately.

Audio-process status no longer performs a WMI/COM scan every second. Normal
monitoring uses the process ID returned at launch, with WMI retained only as a
one-time recovery fallback. This prevents the AutoHotkey memory error observed
after the program was left unattended.

### Controller and interface refinements

- Left/Right now switches between **Keyboard inputs** and **Controller inputs**
  using either keyboard arrows or the controller D-pad.
- Down from **Keyboard inputs** enters the first shortcut field instead of
  skipping the entire shortcut list.
- Both overlay sliders now use the label **Opacity**.
- **Maximum PNG size** is aligned more clearly, and **About...** remains visible
  at the preferred magnetic-snap size.

### LaunchBox / Big Box and documentation

The per-game setup window now includes **Open JRPG Translator...** with concise
first-time guidance. It helps users configure API keys, hotkeys, capture, and
overlay settings before relying on hidden background launches. The window was
resized and rearranged so its final explanatory lines remain visible.

The repository README now includes a visual showcase for screenshot
translation, transcripts with kanji readings, explanations, live audio, and
LaunchBox integration, with an updated plugin screenshot.

## Included source

The repository contains the AutoHotkey control panel and overlays, Python
translation scripts, bundled multilingual prompts, and LaunchBox / Big Box
plugin source. Personal API keys, Profiles, runtime settings, compiled
executables, backups, logs, screenshots, and generated user data are not
included.

For the complete change list, see `CHANGELOG.md`.
