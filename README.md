# JRPG Translator Toolkit

JRPG Translator is a Windows toolkit for translating Japanese games while you
play. It combines screenshot translation, direct live-audio translation, and a
separate Japanese-learning explainer with customizable overlay windows.

The control panel works with a mouse and keyboard or directly from an
XInput-compatible controller. Keyboard mapping tools such as JoyToKey, Steam
Input, or DS4Windows remain optional for custom and multi-function mappings.

## See It in Action

### Live Audio Translation

Translated dialogue appears in the overlay while a voiced cutscene is playing.

<p align="center">
  <a href="docs/media/live-audio-translation.gif">
    <img src="docs/media/live-audio-translation.gif" alt="Live audio translation appearing during a voiced game cutscene" width="800">
  </a>
</p>

### Screenshot Translation Modes

Choose a compact translation-only overlay or include the Japanese transcript
and kanji readings for language study. Click either image to view it at full
size.

<table>
  <tr>
    <th width="50%">Translation only</th>
    <th width="50%">Transcript, kanji readings, and translation</th>
  </tr>
  <tr>
    <td>
      <a href="docs/media/screenshot-translation.jpg">
        <img src="docs/media/screenshot-translation.jpg" alt="Japanese game with a compact English translation overlay">
      </a>
    </td>
    <td>
      <a href="docs/media/transcript-kanji-readings.jpg">
        <img src="docs/media/transcript-kanji-readings.jpg" alt="Japanese game with transcript, kanji readings, and English translation">
      </a>
    </td>
  </tr>
</table>

### Explanations and Frontend Integration

The separate Explainer can break down vocabulary, readings, grammar, and nuance.
The optional LaunchBox / Big Box plugin prepares the selected JRPG Translator
Profile automatically for each game.

<table>
  <tr>
    <th width="50%">Japanese-learning explanation</th>
    <th width="50%">LaunchBox / Big Box integration</th>
  </tr>
  <tr>
    <td>
      <a href="docs/media/japanese-explainer.png">
        <img src="docs/media/japanese-explainer.png" alt="Detailed Japanese language explanation displayed over a game">
      </a>
    </td>
    <td>
      <a href="docs/media/launchbox-integration.jpg">
        <img src="docs/media/launchbox-integration.jpg" alt="Per-game JRPG Translator and JoyToKey setup window in LaunchBox">
      </a>
    </td>
  </tr>
</table>

### Study Library and Reader

Save explanations together with their Japanese source and optional screenshots,
organize them with searchable metadata, and revisit them in a focused reading
view. Click either image to view it at full size.

<table>
  <tr>
    <th width="50%">Study Library</th>
    <th width="50%">Study Reader</th>
  </tr>
  <tr>
    <td>
      <a href="docs/media/study-library.png">
        <img src="docs/media/study-library.png" alt="Searchable Study Library with saved explanations, source screenshots, metadata, and Anki status">
      </a>
    </td>
    <td>
      <a href="docs/media/study-reader.png">
        <img src="docs/media/study-reader.png" alt="Study Reader showing a Japanese explanation beside its source screenshot and context">
      </a>
    </td>
  </tr>
</table>

Review saved sentences and vocabulary after playing, optionally ask the selected
AI model to rank useful study candidates, and send chosen material to Anki after
checking the editable card preview.

<table>
  <tr>
    <th width="50%">Review for Anki</th>
    <th width="50%">Add explanation to Anki</th>
  </tr>
  <tr>
    <td>
      <a href="docs/media/review-for-anki-v095.png">
        <img src="docs/media/review-for-anki-v095.png" alt="Review for Anki window with AI-ranked sentence candidates from the Study Library">
      </a>
    </td>
    <td>
      <a href="docs/media/add-explanation-to-anki-v095.png">
        <img src="docs/media/add-explanation-to-anki-v095.png" alt="Editable Anki card preview with Japanese text, explanation, destination deck, and source screenshot">
      </a>
    </td>
  </tr>
</table>

## Features

- Screenshot translation with OpenAI and Google Gemini vision models.
- Near-live audio translation using OpenAI Realtime Translate or Gemini Live
  Translate models, without a separate transcription step.
- A dedicated Explainer for vocabulary, kanji readings, grammar, literal
  meaning, natural translations, nuance, and cultural context.
- A searchable Study Library and focused Study Reader with optional source
  screenshots, metadata, version history, editing, copying, and Anki status.
- Optional automatic saving of plain-text explanation copies for later study.
- Independent Translator and Explainer overlays with configurable colors,
  fonts, borders, transparency, position, and size.
- Translation and explanation prompts editable from the control panel.
- Unified Profiles that store prompt selections, capture target, terminology
  settings, D-pad navigation, startup choices, and both overlays' appearance,
  position, and size.
- Independent model lists for screenshot translation, live audio, and
  explanations, with API-backed model discovery and manual model-ID entry.
- Selectable output language for live audio translation.
- JP-to-target-language and target-language-to-target-language glossary profiles
  for consistent names, terminology, spelling, and preferred wording, with an
  immediate per-Profile on/off switch.
- Configurable keyboard hotkeys, optional direct controller action bindings,
  optional D-pad navigation, and spatial controller focus movement.
- A resizable control panel with compact scrolling, dark mode, whole-window
  opacity, and magnetic preferred width and height points.
- Non-activating overlays that can remain visible without taking focus from the
  game or pausing an emulator.
- Optional per-game LaunchBox / Big Box integration with JRPG Translator
  Profile selection and independent JoyToKey profile switching.

## Requirements

- Windows 10 or Windows 11.
- An OpenAI API key, a Gemini API key, or both.
- Internet access for translation and explanation requests.

The downloadable release includes a portable Python environment and compiled
AutoHotkey executables. A separate Python or AutoHotkey installation is not
needed when using the release package.

## Quick Start

1. Download and extract the
   [latest release](https://github.com/retrogamer0815/jrpg-translator-toolkit/releases/latest).
2. Run `JRPG Translator.exe`.
3. Add your API keys in the **API Keys** tab, or set `OPENAI_API_KEY` and/or
   `GEMINI_API_KEY` as Windows user environment variables.
4. Open the Translator overlay.
5. In **Screenshot Translation**, choose **Capture...**, then select a region or
   game window.
6. Select the provider, model, prompt, and hotkeys you want to use.

Most control-panel settings save immediately. Manually edited application paths
and **Debug mode** are saved with **Save paths** in the optional **Paths** tab.
API keys stored in `Settings/.env` use **Save Keys** in **API Keys**; prompt and
terminology editors have their own Save controls.

## Translation Workflows

### Screenshot Translation

Use **Screenshot + Translate** for the fastest one-button workflow. You can also
take one or more screenshots first and translate them together, which is useful
when a sentence spans multiple dialogue boxes.

The screenshot prompt controls the requested output, so it can include a plain
translation, the original Japanese, kanji readings, speaker names, or any other
format useful for playing or studying.

### Live Audio Translation

In the **Audio Translation** tab:

1. Select the Windows playback device.
2. Optionally choose **Test Audio** while sound is playing to verify that the
   selected device is receiving audible output. This test is local and does not
   make an API request.
3. Choose OpenAI or Gemini and a compatible live translation model.
4. Select the output language.
5. Choose **Audio Translation Off** in the footer to start translation. The
   button changes to **Audio Translation On** while the live session is active;
   choose it again to stop.

Audio is streamed directly to the selected live translation model. Translated
lines appear at the bottom of the Translator overlay while older lines move
upward and remain available for scrolling.

### Japanese Explainer

The Explainer uses the most recent Japanese screenshot text. Choose
**Explain last jp. Text** or its configured hotkey to generate a separate
learning-focused explanation without replacing the translation.

Explanation prompts are independent from translation prompts, so they can be
tuned for a learner's level and preferred amount of detail.

### Study Library and Reader

Enable **Save explanations to Study Library** to keep the Japanese source,
explanation, Profile, provider/model details, and optional source screenshots in
a local searchable database. Repeated explanations of the same Japanese text
are grouped as versions rather than becoming unrelated entries.

**Open Study Library...** provides filters for Profile, chapter, speaker, tags,
date/time, and Anki status, plus bulk metadata editing and configurable columns.
Its **Study** action opens a reading-focused view with the explanation, parsed
section navigation, screenshots, and original Japanese. Explanations can be
edited safely or copied—including an option that removes attached hiragana
readings from the Japanese text before it is pasted into an Anki card.

Multiple named Study Libraries can be created and selected through unified
Profiles. Optional **Save plain-text copies** continues to use
`Settings/Explanations`, sorted into a subfolder for the active Profile.

### Terminology Overrides

The **Terminology Overrides** tab provides separate JP-to-target-language and
target-language-to-target-language glossary Profiles. **Use terminology
overrides** enables or disables both stages immediately for screenshot
translation, live audio translation, and explanations.

When the option is off, JP-to-target-language entries are not sent to a model
and local target-language replacements are skipped. The enabled state and both
selected glossary Profiles are stored in unified Profiles, including Profiles
applied through LaunchBox / Big Box.

Each glossary type has an independent profile list. **Manage Entries...** opens
a two-column table where mappings can be added, edited, or deleted without
manually maintaining `source -> replacement` lines. TL -> TL entries are applied
locally to correct model output; JP -> TL entries are additional instructions
sent to the selected model.

## Controller Use

The control panel supports an XInput-compatible controller without a keyboard
mapper:

- The D-pad moves spatially between visible controls and tabs.
- A / Cross confirms; B / Circle cancels or closes dialogs and text editors.
- **Use D-pad for control panel navigation** can be turned off when JoyToKey,
  Steam Input, DS4Windows, or another mapper already sends arrow keys. A / Cross
  and B / Circle remain available.
- The **Controls** tab can optionally bind actions such as Screenshot +
  Translate, Explain, and overlay visibility directly to controller buttons.
- Keyboard hotkeys remain available and can still be mapped through JoyToKey or
  another controller mapper for custom, long-press, and multi-function layouts.
- Font size, Max PNG size, transparency, font weight, and overlay colors can be
  adjusted without a mouse. The controller color editor provides live hue,
  saturation, and brightness previews.
- Mouse-wheel or arrow-key mappings can scroll the visible Translator or
  Explainer overlay even when it does not own game focus.

The overlays can be brought forward without becoming the active window. This
allows emulator options such as RetroArch's pause-when-inactive behavior to
remain enabled while translations are visible. Opening the control panel still
activates it normally.

The Translation Window and Explanation Window tabs also include a
`Move / Resize` mode. With an XInput-compatible controller, the left stick moves
the selected overlay and the right stick resizes it. Arrow keys provide a
fallback; hold the configured Screenshot + Translate key while pressing arrows
to resize. Enter saves the new bounds and Escape restores the previous bounds.

Capture targets can also be configured without a mouse. **Capture > Region**
reuses the analog-stick move and resize controls, while **Capture > Window**
cycles through available windows and previews the selected target. A / Cross
saves and B / Circle cancels either mode. Opening Region mode with a mouse keeps
the conventional drag-to-select workflow.

## LaunchBox / Big Box Integration

A preview plugin is included as source under `integrations/launchbox`. It adds
`JRPG Translator Setup...` to each game's LaunchBox context menu and Big Box
details menu. Per game, it can:

- start JRPG Translator in background mode and close only the instance it
  started;
- apply a selected JRPG Translator Profile for the game;
- start JoyToKey or switch an existing instance to a selected profile; and
- restore the previous JoyToKey profile when the game exits.

If JRPG Translator is already running, the plugin applies the selected Profile
without restarting it and leaves the pre-existing instance open when the game
exits.

The plugin setup window can browse for the Translator executable, JoyToKey
executable, and JoyToKey profile folder. Big Box uses a controller-native path
browser, while LaunchBox retains the standard Windows file and folder pickers.
The setup window also provides **Open JRPG Translator...** for first-time API,
hotkey, capture, and overlay configuration before using background launches.
See
[`integrations/launchbox/README.md`](integrations/launchbox/README.md) for build,
packaging, and installation instructions.

## API Keys and Privacy

The recommended key-storage method is Windows user environment variables:

```text
OPENAI_API_KEY=your_key
GEMINI_API_KEY=your_key
```

The control panel can alternatively store keys in `Settings/.env`. This is a
plain-text file: do not commit it, upload it, or include it in shared archives.

The **API Keys** tab can open Windows Environment Variables directly. If a
selected provider has no configured key, the relevant overlay identifies the
missing OpenAI or Gemini key instead of leaving a loading indicator active.

Screenshots and audio sent for translation are processed by the selected API
provider. Review the provider's current data and privacy terms before use.

## Settings and Profiles

### Desktop layout

The desktop control center now opens with grouped sidebar navigation and a compact profile shortcut. **Screenshot Translation** is the first fully refreshed page: three softly rounded panels separate AI settings, capture actions, and formatting. Only the active provider's model is shown. **Capture & translate** is the primary action; model/prompt management use lighter link-style buttons. Multi-screenshot guidance and **Startup & advanced capture** expand when needed. The AI fields stack in narrower windows; content width is capped on wide displays and long pages scroll between the fixed header and footer.

**Audio Translation** follows the same design, with panels for **AI settings**, **Audio input**, and **Live translation**. Provider, active model, and output language share a row when space permits. **Manage models…** replaces the separate provider-specific Add/Delete rows. The device selector, **Refresh devices**, and **Test audio** are grouped with an **Input check** readout; longer testing/troubleshooting help expands on demand. A prominent **Start audio translation** / **Stop audio translation** button uses the existing audio-session workflow. Device checks remain local and asynchronous, with the existing 30-second timeout. Current selections, live status, and the original classic audio form are preserved.

**Explanation** now has matching panels for **AI settings**, **Create an explanation**, and **Save explanations**. Its provider, active model, and prompt stay independent of Screenshot Translation. **Explain latest text** uses the existing explanation action, with **Open Study Library** beside it. Library saving, source screenshots, and independent plain-text copies have short inline guidance; screenshot saving is available only while Library saving is enabled. **Startup options** expands to show Explainer startup and always-on-top preferences. Existing settings, callbacks, and the classic form are preserved.

**Overlay windows** opens directly to the last-used **Translator** or **Explainer** settings page, with selectors to switch between them. Both use matching **Appearance**, **Typography**, and **Position & size** panels. Color controls show a swatch and hexadecimal value; the Translator also retains its separate speaker-name color. Each overlay keeps its own opacity, colors, font, size, and bold setting. The existing color picker and **Move / resize…** workflow are unchanged. Opacity still affects the whole overlay, including its text.

**Profiles** groups **Saved profiles**, **Startup overlays**, and **What a profile includes** into matching cards. Choosing a profile only selects it; **Apply profile** loads it, while **Save current** updates it with the current setup. The page distinguishes saved prompts, capture/formatting, terminology, overlay preferences, D-pad navigation, Library selection and configured Anki mappings from global AI models/providers, audio input/language, keyboard shortcuts and controller action bindings. Existing profile files and save/apply/delete behavior are unchanged.

**Settings** opens directly to its last-used section: **Controls**, **Terminology**, **API keys**, or **Paths** (only when enabled). Controls offers Keyboard/Controller selectors, responsive binding rows, duplicate-shortcut emphasis and wrapped conflict messages, with controller navigation options in a separate card. Terminology separates local output corrections from instructions sent to the model. API keys separates Windows environment variables from optional plain-text `.env` storage, preserving masked inputs and existing save/enabled states. Paths groups executable/script locations and advanced diagnostic options. **Window options…** and the reversible **Use classic layout** remain available. Existing dialogs and configuration callbacks are reused.

The footer has one control each for the Translator, Explainer, and audio session. **Window options…** contains the desktop light/dark preference, always-on-top option, and opacity. New installations default to dark; an existing appearance preference is preserved. Fullscreen retains its separate opacity and always-on-top settings.

The modern desktop uses one integrated header with the app name, profile shortcut, and custom minimize/maximize/close controls. Drag an empty part of the header (or the app name) to move it; double-click to maximize or restore. Native edge resizing and window keyboard shortcuts remain available. Dropdowns throughout the refreshed main pages use rounded, subdued borders with an integrated chevron and a clear keyboard-focus outline; their native selection and popup behavior are preserved. The classic fallback restores the standard title bar and dropdown styling.

The desktop header includes the supplied JRPG Translator logo beside the app name. Its visible artwork fits a 52-pixel box at 100% scaling inside a 72-pixel header, with a 14-point title and a 20-pixel gap; all scale with display DPI. Rendering ignores empty transparent margins while preserving the original PNG; the image is embedded in the executable so the original source file is not required. The logo remains part of the draggable header and does not add a keyboard focus stop. Fullscreen branding is unchanged.

The profile shortcut opens Profiles; it does not apply a profile automatically. Switching layouts does not change current model, prompt, capture, overlay, profile, or other configuration values. The layout preference is stored as `modernLayout` in `[cfg_control]`. Fullscreen retains its separate presentation.

Desktop **Study Library** and **Study Reader** now use the same integrated logo header, rounded panels, quieter buttons and dark dropdown/header styling. The Library separates search/filter controls from actions on selected explanations; library management stays beside the library selector, and storage/export remain in the footer. The Reader keeps version/section navigation above a borderless reading area, with screenshots, Japanese source and scrollable metadata in a context panel. Entry navigation and the manual Anki marker stay in its footer. Native text selection, editing, columns, sorting and Anki workflows are preserved. These windows resize independently (minimum 960 × 700 logical pixels), and their text/table fields scroll internally without moving the header/footer. Newly opened Study windows use the classic presentation when the classic layout is selected; Big Box is unchanged.

Desktop **Add to Anki** card review and **Review for Anki** now share that design. Card review places editable front/back text beside the destination deck, profile mapping and fitted screenshot; **Generate example…** remains available for vocabulary. Review for Anki has grouped filters, clearly selected Sentences/Vocabulary tabs, and assessment details below its table. Confirmations and notices owned by modern Study windows use the integrated header and larger plain text; long messages remain scrollable. The underlying review, confirmation, duplicate detection and creation workflow is unchanged. Classic and fullscreen dialogs retain their existing presentation.

Desktop **Current chapter**, **Columns**, **Filters**, and single-explanation **Edit details** use the same integrated header, grouped dark panels and clear primary actions. Chapter history remains editable, required Japanese-source visibility stays protected, and filter dates use matching dark date/time pickers. Confirming a date only updates the dialog's draft; **Apply filters** still validates and applies the whole selection. Fullscreen and classic layouts remain unchanged.

Desktop **Manage Study Libraries**, **Archived libraries**, and **Library storage** also use the modern shell. Library tables have readable column widths and separate library-wide and selected-row actions. New, Rename, Restore, and their confirmations share the same styling; Default-library protections and existing data workflows are preserved. Storage groups usage figures into two columns and wraps long folder paths in a selectable, read-only field. Fullscreen and classic presentations remain available unchanged.

Desktop **Generate/Regenerate recommendations**, **Recommendation preferences**, and the **Recommendation prompt** editor share the same resizable panels, dropdowns and integrated header. Candidate/model details are separate from selection options; preferences group study focus and optional guidance. **Save draft** keeps prompt edits in the preferences dialog, and **Apply preferences** accepts them. The full-prompt view is explicitly read-only and preserves the editable draft when switching back. Protected output rules, recommendation caching, and explicit Generate authorization are unchanged; opening or previewing these dialogs makes no AI request. Fullscreen and classic layouts are preserved.

Desktop **Anki connection and link check**, **Bulk edit**, and the Reader's **Generate new explanation version** now use the same modern dialog shell. Connection status is separate from profile/deck/field mapping; bulk changes still start at **Keep existing**; new-version AI settings show only the selected provider's model. The Reader's **Edit prompt** and **New prompt** child dialogs now have modern desktop and dedicated fullscreen presentations. Prompt editing retains the saved-file/backup behavior, and creating a prompt does not generate an explanation. Child dialogs restore the parent's enabled state and focus on close. Classic layouts and the existing fullscreen connection, bulk-edit and generation forms remain available.

**Shared messages and action menus (batch 2)** follow their actual owner window, including nested dialogs, rather than a global presentation-mode flag. Calls through the shared confirmation/notice helpers use the modern desktop shell or the existing fullscreen message presentation, retaining No as the safe confirmation default and scrollable long details. This includes Reader prompt/generation/edit notices, Study Library/Anki errors and export results; bulk-edit validation now uses the same route. Desktop **Anki**, Reader **Copy/Edit**, and Library/Review right-click menus use quieter borders, roomier rows, hover/keyboard highlighting and monitor work-area positioning. Their fullscreen counterparts use large action buttons, Back to cancel, and pages of up to six choices. Disabled items remain unavailable and returned action indices are unchanged. This batch does not replace remaining direct native `MsgBox` calls, generic name inputs, file/folder pickers, or other unmodernized editors.

**Name-entry dialogs (batch 3)** now use the integrated desktop header, a clean text field and consistent OK/Cancel actions. This covers new game profiles, translation/explanation prompt names, both terminology-profile names and manual model IDs. The shared input helper also provides a dedicated fullscreen form when called from a fullscreen owner; existing in-page fullscreen workflows are unchanged. Blank/duplicate-name notices and overwrite/delete confirmations in these creation/management actions use the shared modern messages. Cancelling returns no draft value, parent enabled state/focus is restored, and the existing name validation and file operations stay in their callers. Larger prompt/terminology editors, the online model browser and unrelated native messages remain for later batches.

**Model-selection dialogs (batch 4)** modernize the Add model source chooser and online model browser for Screenshot Translation, Audio Translation and Explanation. The chooser keeps Browse online / Enter model ID separate from the explicit Continue action. The browser separates compatible models from catalog/cache status, scrolls long IDs, skips disabled actions when the list is empty, and retains the existing duplicate filter and refresh behavior. Related catalog-failure and missing-selection notices use shared messages. These dialogs follow their desktop or fullscreen owner and restore its enabled state/focus when closed; the existing Big Box in-page model management and classic layouts are unchanged. Browsing a catalog may query the provider's model-list endpoint, but does not generate content or add a model until confirmed. Larger prompt/terminology editors, controls dialogs and unrelated native messages remain for later batches.

Study-only layout checks: `pwsh -NoProfile -File tests/test_desktop_layout.ps1 -StudyOnly`. They construct the real Study controls with temporary settings, verify compact/normal/wide layouts, button text fit, screenshot bounds, editing controls and cleanup, and render synthetic screenshots without reading a personal library or contacting Anki. Use `-DialogsOnly` for the Anki, metadata, management, recommendation, shared message/menu and name-entry checks. The full run also checks shutdown with modern Study/Anki/name-entry windows left open, including restoration of native table-header procedures and detachment of date-picker callbacks before GUI teardown.

Desktop page scrolling uses native mouse-wheel messages, coalesced into frames, rather than global wheel hotkeys. Rapid scrolling therefore does not consume AutoHotkey's hotkey flood limit; open dropdowns and opacity sliders retain their native wheel behavior. Page switches focus the selected section, and scrolling clips controls before moving them so text cannot briefly spill into the fixed header or footer.

On exit, pending scroll frames are cancelled and native wheel handlers are detached before GUI/global cleanup, preventing late destruction callbacks from opening an unassigned-variable error dialog. The desktop test suite includes an isolated live-window exit regression, normal parent/child destruction, and late callbacks with their tracking globals already released. Run only these checks with `pwsh -NoProfile -File tests/test_desktop_layout.ps1 -ShutdownOnly`.

The isolated desktop layout checks can be run with `pwsh -NoProfile -File tests/test_desktop_layout.ps1`; they exercise production layout functions against synthetic controls and temporary settings, without launching overlays or sending API requests. Audio coverage includes responsive fields, native dropdowns, tab order, inline diagnostic states, unclipped expanded help, session-state labels, and preservation of selections through page and classic-layout round trips. Explanation coverage adds independent AI choices, the original action callback, dependent Library/screenshot preferences, separate plain-text saving, expandable startup settings, and classic restoration. Overlay coverage checks both selectors, color values and targets, independent font/opacity settings, native size ranges, keyboard navigation, narrow layouts, move/resize callback wiring, and classic restoration. Profiles/Settings coverage includes profile selection without applying, retained callbacks and values, both input views, duplicate-shortcut warnings, masked API keys and enabled states, unsaved path edits, hidden Paths handling, and all classic-layout round trips. Scroll regressions cover wheel-style steps and coalesced scrollbar dragging on all ten main pages, with opaque and translucent windows: the main window must never suspend its redraw state, and fixed header/sidebar/footer controls must not repaint. Modern desktop painting is buffered and invalidation is limited to the content viewport; classic buffering and fullscreen behavior remain unchanged.

Additional input regressions deliver 240-message native wheel bursts in both directions to parent and dropdown windows, exercise fractional wheel deltas and page changes with a queued scroll frame, verify selected-section focus, and inspect clipping regions during native move callbacks. A deliberately unprotected move is used as a positive control; completed-frame pixel comparisons alone would miss this transient clipping defect.

**Profiles → Startup overlays** (desktop and full-screen) offers **None**,
**Translator only**, **Explainer only**, or **Translator and Explainer**.
Apply a Profile first if you want to edit its setup, choose the startup overlays,
then use **Save Current** / **Save current settings** to store the choice in that
Profile. The existing startup checkboxes stay synchronized. The LaunchBox / Big
Box plugin respects these settings on a fresh tool start; it no longer forces
the Translator overlay to open. Running overlays are not opened or closed when
changing startup settings or applying a Profile to an already-running tool.

Portable settings live in the local `Settings` folder:

```text
Settings/
|-- control.ini
|-- .env                         # optional local API-key storage
|-- Screenshots/
|-- Explanations/                # optional saved study material
|-- prompts/                     # screenshot translation prompts
|-- prompts_explain/             # explanation prompts
|-- glossaries/
`-- game_profiles/               # unified Profiles
```

A unified Profile stores the selected screenshot and explanation prompts,
capture region or window, guessed-subject highlighting, speaker-name coloring,
terminology selections and enabled state, D-pad navigation, overlay startup and
topmost choices, and both overlays' position, size, colors, transparency, font,
font size, and weight. Translator and Explainer appearance settings remain
independent inside each Profile. Screenshot output handling is selected
automatically from the prompt name and is no longer a normal user-facing
selector.

The source repository and release include three prompt families for every
supported output language:

| Prompt pattern | Purpose |
| --- | --- |
| `Settings/prompts/default_<language>.txt` | Screenshot translation with a plain Japanese transcript |
| `Settings/prompts/default_with_kanji_reading_<language>.txt` | Screenshot translation with hiragana readings added to kanji words |
| `Settings/prompts_explain/default_<language>.txt` | Japanese-learning explanation in the selected language |

Available language labels are `en`, `de`, `fr`, `es`, `it`, `pt`, `nl`, `pl`,
`ru`, `uk`, `ko`, `zh-CN`, `zh-TW`, and `ja`. The screenshot prompts retain the
literal `Transcript:` and `Translation:` headings because the output parser uses
them; the translated content follows the language requested by the selected
prompt. When several screenshots are submitted together, the bundled prompts
instruct the model to reconstruct the passage in capture order before
translating it.

Additional prompt profiles created through the control panel remain local and
are ignored by Git.

## Project Structure

| File | Purpose |
| --- | --- |
| `JRPG Translator.ahk` | Main control panel and workflow orchestration |
| `bin/overlay.ahk` | Translator and Explainer overlay windows |
| `bin/overlay.exe` | Compiled overlay included in release packages, not the source repository |
| `scripts/screenshot_translator.py` | Screenshot vision translation and output formatting |
| `scripts/live_audio_translator.py` | Direct streaming audio translation |
| `scripts/explainer.py` | Japanese-learning explanations |
| `scripts/model_catalog.py` | Provider model discovery, filtering, sorting, and caching |
| `integrations/launchbox/` | Optional per-game LaunchBox / Big Box and JoyToKey integration |
| `docs/media/` | Curated screenshots and animation displayed in this README |

Runtime messages and generated overlay text are exchanged through
`%TEMP%\JRPG_Overlay`.

## Running from Source

Install AutoHotkey v2 and run `JRPG Translator.ahk`. The source version launches
`bin/overlay.ahk`; compiled releases launch `bin/overlay.exe`.
The Python scripts require Python 3.12 and their listed dependencies, or the
portable Python environment included in a release package.

```powershell
py -3.12 -m pip install -r requirements.txt
```

Before sharing a build, verify that it does not contain `Settings/.env`, API
credentials, personal Profiles, `Settings/Screenshots`, saved explanations,
logs, or other local state. The curated showcase files under `docs/media` are
intentional repository assets.

## Troubleshooting

- If an overlay reports a missing API key, add it through **API Keys** or set the
  provider's Windows user environment variable and restart JRPG Translator.
- If a request fails, verify the selected model name and confirm that the API
  key has access to that model.
- If the wrong playback source is translated, refresh and reselect the device
  in **Audio Translation**, then use **Test Audio** while sound is playing.
- If an overlay is missing, use the Open Translator or Open Explainer button and
  check its saved position on connected displays.
- If source files do not start, confirm that AutoHotkey v2 is being used rather
  than AutoHotkey v1.

## Credits

- [AutoHotkey](https://www.autohotkey.com/): GNU GPLv2.
- [Python](https://www.python.org/): PSF License.
- [PixelMplus](https://itouhiro.github.io/mplus-fonts/): SIL Open Font License
  1.1.
- Application icon by Miguel C Balandrano via Flaticon; attribution is required
  by the source license.

## License

The project source is released under the MIT License. See [LICENSE](LICENSE).
