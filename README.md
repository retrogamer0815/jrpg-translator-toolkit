# JRPG Translator Toolkit

JRPG Translator is a Windows toolkit for translating Japanese games while you
play. It combines capture-based game text translation, direct live-audio translation, and a
separate Japanese-learning explainer with customizable overlay windows.

The control panel works with a mouse and keyboard or directly from an
XInput-compatible controller. Keyboard mapping tools such as JoyToKey, Steam
Input, or DS4Windows remain optional for custom and multi-function mappings.

## User manual

Read the **[complete user manual for v0.9.9 (in testing)](docs/manual/README.md)**
for setup, controllers, every main feature, LaunchBox / Big Box, Study Library,
Anki, troubleshooting, and backups. The manual includes real screenshots, a labeled
desktop overview, and clearly marked placeholders for the remaining screenshots.

New users can start with **[Setup and your first translation](docs/manual/02-setup.md)**.
The manual describes the current interface; older screenshots and development notes
elsewhere in this README may show previous versions.

## See It in Action

### Live Audio Translation

Translated dialogue appears in the overlay while a voiced cutscene is playing.

<p align="center">
  <a href="docs/media/live-audio-translation.gif">
    <img src="docs/media/live-audio-translation.gif" alt="Live audio translation appearing during a voiced game cutscene" width="800">
  </a>
</p>

### Game Text Translation Modes

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

- Game text translation from captures with OpenAI and Google Gemini vision models.
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
- Independent model lists for Game Text Translation, live audio, and
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
5. In **Game Text Translation**, choose **Capture...**, then select a region or
   game window.
6. Select the provider, model, prompt, and hotkeys you want to use.

Most control-panel settings save immediately. API keys stored in `Settings/.env`
use **Save Keys** in **API Keys**; prompt and terminology editors have their own
Save controls.

### Advanced configuration

The normal desktop and fullscreen interfaces do not expose runtime paths or
diagnostic switches. Advanced users can edit these values in
`Settings/control.ini` under `[cfg]`, then restart JRPG Translator:

```ini
[cfg]
pythonExe=.\python\python.exe
directModelOutput=0
debugMode=0
```

`pythonExe` accepts an absolute path or a path relative to the application
folder. `directModelOutput` and `debugMode` use `0` for off and `1` for on; both
default to off.

## Translation Workflows

### Game Text Translation

Use **Capture + Translate** for the fastest one-button workflow. You can also
make one or more captures first and translate them together, which is useful
when a sentence spans multiple dialogue boxes.

The Game Text prompt controls the requested output, so it can include a plain
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
- The **Controls** tab can optionally bind actions such as Capture +
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
fallback; hold the configured Capture + Translate key while pressing arrows
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

The desktop control center now opens with grouped sidebar navigation and a compact profile shortcut. **Game Text Translation** is the first fully refreshed page: three softly rounded panels separate AI settings, capture actions, and formatting. Only the active provider's model is shown. **Capture & Translate** is the primary action; model/prompt management use lighter link-style buttons. Multi-capture guidance and **Startup & advanced capture** expand when needed. The AI fields stack in narrower windows; content width is capped on wide displays and long pages scroll between the fixed header and footer.

**Audio Translation** follows the same design, with panels for **AI settings**, **Audio input**, and **Live translation**. Provider, active model, and output language share a row when space permits. **Manage models…** replaces the separate provider-specific Add/Delete rows. The device selector, **Refresh devices**, and **Test audio** are grouped with an **Input check** readout; longer testing/troubleshooting help expands on demand. A prominent **Start audio translation** / **Stop audio translation** button uses the existing audio-session workflow. Device checks remain local and asynchronous, with the existing 30-second timeout. Current selections, live status, and the original control callbacks are preserved.

**Explanation** now has matching panels for **AI settings**, **Create an explanation**, and **Save explanations**. Its provider, active model, and prompt stay independent of Game Text Translation. **Explain latest text** uses the existing explanation action, with **Open Study Library** beside it. Library saving, source screenshots, and independent plain-text copies have short inline guidance; screenshot saving is available only while Library saving is enabled. **Startup options** expands to show Explainer startup and always-on-top preferences. Existing settings and callbacks are preserved.

**Overlay windows** opens directly to the last-used **Translator** or **Explainer** settings page, with selectors to switch between them. Both use matching **Appearance**, **Typography**, and **Position & size** panels. Color controls show a swatch and hexadecimal value; the Translator also retains its separate speaker-name color. Each overlay keeps its own opacity, colors, font, size, and bold setting. The existing color picker and **Move / resize…** workflow are unchanged. Opacity still affects the whole overlay, including its text.

**Profiles** groups **Saved profiles**, **Startup overlays**, and **What a profile includes** into matching cards. Choosing a profile only selects it; **Apply profile** loads it, while **Save current** updates it with the current setup. The page distinguishes saved prompts, capture/formatting, terminology, overlay preferences, D-pad navigation, Library selection and configured Anki mappings from global AI models/providers, audio input/language, keyboard shortcuts and controller action bindings. Existing profile files and save/apply/delete behavior are unchanged.

**Settings** opens directly to its last-used section: **Controls**, **Terminology**, or **API keys**. Controls offers Keyboard/Controller selectors, responsive binding rows, duplicate-shortcut emphasis and wrapped conflict messages, with controller navigation options in a separate card. Terminology separates local output corrections from instructions sent to the model. API keys separates Windows environment variables from optional plain-text `.env` storage, preserving masked inputs and existing save/enabled states. Advanced runtime options remain available in `Settings/control.ini`. **Window options…** remains available; the obsolete classic-layout selector has been removed. Existing dialogs and configuration callbacks are reused.

The footer has one control each for the Translator, Explainer, and audio session. **Window options…** contains the desktop light/dark preference, always-on-top option, and opacity. New installations default to dark; an existing appearance preference is preserved. Fullscreen retains its separate opacity and always-on-top settings.

The desktop uses one integrated header with the app name, profile shortcut, and custom minimize/maximize/close controls. Drag an empty part of the header (or the app name) to move it; double-click to maximize or restore. Native edge resizing and window keyboard shortcuts remain available. Dropdowns throughout the refreshed main pages use rounded, subdued borders with an integrated chevron and a clear keyboard-focus outline; their native selection and popup behavior are preserved.

The desktop header includes the supplied JRPG Translator logo beside the app name. Its visible artwork fits a 52-pixel box at 100% scaling inside a 72-pixel header, with a 14-point title and a 20-pixel gap; all scale with display DPI. Rendering ignores empty transparent margins while preserving the original PNG; the image is embedded in the executable so the original source file is not required. The logo remains part of the draggable header and does not add a keyboard focus stop. Fullscreen branding is unchanged.

The profile shortcut opens Profiles; it does not apply a profile automatically. Desktop now always uses the modern layout. Older settings with `modernLayout=0` in `[cfg_control]` are migrated to `1` on startup; current model, prompt, capture, overlay, profile, and other configuration values are unaffected. Fullscreen retains its separate presentation.

Desktop **Study Library** and **Study Reader** now use the same integrated logo header, rounded panels, quieter buttons and dark dropdown/header styling. The Library separates search/filter controls from actions on selected explanations; library management stays beside the library selector, and storage/export remain in the footer. Clicking a saved-explanation row uses the exact pointer row to refresh the version, screenshot, Japanese source and explanation detail immediately, while arrow-key and D-pad navigation continue to follow native item focus. The Reader keeps version/section navigation above a borderless reading area, with screenshots, Japanese source and scrollable metadata in a context panel. Entry navigation and the manual Anki marker stay in its footer. Native text selection, editing, columns, sorting and Anki workflows are preserved. These windows resize independently (minimum 960 × 700 logical pixels), and their text/table fields scroll internally without moving the header/footer. Desktop Study windows always use the modern presentation; Big Box is unchanged.

Desktop **Add to Anki** card review and **Review for Anki** now share that design. Card review places editable front/back text beside the destination deck, profile mapping and fitted screenshot; **Generate example…** remains available for vocabulary. Review for Anki has grouped filters, clearly selected Sentences/Vocabulary tabs, and assessment details below its table. Confirmations and notices owned by modern Study windows use the integrated header and larger plain text; long messages remain scrollable. The underlying review, confirmation, duplicate detection and creation workflow is unchanged. Fullscreen dialogs retain their existing presentation.

Desktop **Current chapter**, **Columns**, **Filters**, and single-explanation **Edit details** use the same integrated header, grouped dark panels and clear primary actions. Chapter history remains editable, required Japanese-source visibility stays protected, and filter dates use matching dark date/time pickers. Confirming a date only updates the dialog's draft; **Apply filters** still validates and applies the whole selection. Fullscreen remains unchanged.

Desktop **Manage Study Libraries**, **Archived libraries**, and **Library storage** also use the modern shell. Library tables have readable column widths and separate library-wide and selected-row actions. New, Rename, Restore, and their confirmations share the same styling; Default-library protections and existing data workflows are preserved. Storage groups usage figures into two columns and wraps long folder paths in a selectable, read-only field. Fullscreen remains available unchanged.

Desktop **Generate/Regenerate recommendations**, **Recommendation preferences**, and the **Recommendation prompt** editor share the same resizable panels, dropdowns and integrated header. Candidate/model details are separate from selection options; preferences group study focus and optional guidance. **Save draft** keeps prompt edits in the preferences dialog, and **Apply preferences** accepts them. The full-prompt view is explicitly read-only and preserves the editable draft when switching back. Protected output rules, recommendation caching, and explicit Generate authorization are unchanged; opening or previewing these dialogs makes no AI request. Fullscreen is preserved.

Desktop **Anki connection and link check**, **Bulk edit**, and the Reader's **Generate new explanation version** now use the same modern dialog shell. Connection status is separate from profile/deck/field mapping; bulk changes still start at **Keep existing**; new-version AI settings show only the selected provider's model. The Reader's **Edit prompt** and **New prompt** child dialogs now have modern desktop and dedicated fullscreen presentations. Prompt editing retains the saved-file/backup behavior, and creating a prompt does not generate an explanation. Child dialogs restore the parent's enabled state and focus on close. The existing fullscreen connection, bulk-edit and generation forms remain available.

**Shared messages and action menus (batch 2)** follow their actual owner window, including nested dialogs, rather than a global presentation-mode flag. Calls through the shared confirmation/notice helpers use the modern desktop shell or the existing fullscreen message presentation, retaining No as the safe confirmation default and scrollable long details. This includes Reader prompt/generation/edit notices, Study Library/Anki errors and export results; bulk-edit validation now uses the same route. Desktop **Anki**, Reader **Copy/Edit**, and Library/Review right-click menus use quieter borders, roomier rows, hover/keyboard highlighting and monitor work-area positioning. Their fullscreen counterparts use large action buttons, Back to cancel, and pages of up to six choices. Disabled items remain unavailable and returned action indices are unchanged. This batch does not replace remaining direct native `MsgBox` calls, generic name inputs, file/folder pickers, or other unmodernized editors.

**Name-entry dialogs (batch 3)** now use the integrated desktop header, a clean text field and consistent OK/Cancel actions. This covers new game profiles, translation/explanation prompt names, both terminology-profile names and manual model IDs. The shared input helper also provides a dedicated fullscreen form when called from a fullscreen owner; existing in-page fullscreen workflows are unchanged. Blank/duplicate-name notices and overwrite/delete confirmations in these creation/management actions use the shared modern messages. Cancelling returns no draft value, parent enabled state/focus is restored, and the existing name validation and file operations stay in their callers. Larger prompt/terminology editors, the online model browser and unrelated native messages remain for later batches.

**Model-selection dialogs (batch 4)** modernize the Add model source chooser and online model browser for Game Text Translation, Audio Translation and Explanation. The chooser keeps Browse online / Enter model ID separate from the explicit Continue action. The browser separates compatible models from catalog/cache status, scrolls long IDs, skips disabled actions when the list is empty, and retains the existing duplicate filter and refresh behavior. Related catalog-failure and missing-selection notices use shared messages. These dialogs follow their desktop or fullscreen owner and restore its enabled state/focus when closed; the existing Big Box in-page model management is unchanged. Browsing a catalog may query the provider's model-list endpoint, but does not generate content or add a model until confirmed. Larger prompt/terminology editors, controls dialogs and unrelated native messages remain for later batches.

**Prompt and terminology editors (batch 5)** give screenshot and explanation prompt files a shared, resizable desktop editor with an integrated header, clear save guidance and the existing atomic-save backup behavior. Model instructions and local corrections use matching terminology tables, add/edit forms, and modern confirmations; malformed files can still be repaired in the raw text editor without losing their original contents. The established fullscreen in-page editors remain unchanged, and opening an editor never sends a model request.

**Control-assignment dialogs (batch 6)** modernize desktop keyboard-shortcut capture and live controller-button assignment. The keyboard form keeps native combination capture behind a readable dark field; controller capture reports disconnected, arming, pressed and release states without changing a binding until the press is complete. Duplicate-shortcut and controller-reassignment prompts now use the shared modern confirmation route with No as the safe default. The existing fullscreen in-page controls workflow remains unchanged.

**Appearance dialogs (batch 7)** bring desktop window preferences and overlay color adjustment into the same resizable shell. Theme, always-on-top and whole-window opacity choices still apply immediately, while the color editor preserves its exact HSV preview, gradient sliders, keyboard/controller navigation and Apply/Cancel behavior. Fullscreen keeps its existing in-page main-window and overlay appearance controls.

**Welcome and About dialogs (batch 8)** give first-run setup guidance, version details, diagnostics, and project/help links the integrated resizable desktop presentation. The welcome checklist retains its dismissal preference and direct route to API Keys; About retains the Welcome Guide, bug-report, GitHub, and copy-version actions. Link, environment-variable, and clipboard failures now follow the owning modern message presentation. Fullscreen continues to use its established in-page About/help screen.

**Operational messages (batch 9)** route the remaining Control Panel notices and decisions through their actual modern desktop or fullscreen owner. This covers explanation/audio helper failures, missing overlay paths, move/resize failures, profile save/apply warnings, API-key storage errors, and unexpected Control Panel errors. System file/folder pickers remain native.

**Picker polish (batch 10)** completes the modern-dialog pass. Desktop **Capture…** now uses the shared compact action menu for Region, Window, and Cancel while retaining near-pointer placement and keyboard/controller activation behavior. The screenshot-folder selector and Study Library workbook export keep their native Windows Explorer interface with explicit app ownership, fullscreen topmost suspension, and enabled-state/focus restoration—even when selection is cancelled or fails.

**Desktop action-menu and focus polish (batch 11)** replaces the remaining native white model/prompt menus on Game Text Translation, Audio Translation, and Explanation with the shared theme-aware action menu. Its roomy rows, selection state, placement, and colors now follow the chosen light or dark desktop theme. Link-style actions no longer retain a blue outline after pointer activation; keyboard/controller navigation still shows an accessible focus cue, with every rounded border edge kept inside the control at high or fractional DPI. The former classic-only native menus are removed.

**Keyboard-shortcut feedback polish (batch 12)** keeps the modern desktop capture field in the selected light or dark theme from its first visible frame, including an unassigned shortcut shown as **None**. Native Windows key-combination capture remains active behind the app-painted field. Saving, removing, restoring a default, or reverting shortcuts now reports the matching result inside the Controls page instead of using a top-right system tooltip; removal is identified explicitly as **Keyboard shortcut removed**. The status clears automatically and restores the normal page description. The existing fullscreen in-dashboard feedback remains unchanged.

Modern desktop keyboard/controller binding displays and in-app API-key inputs use the same restrained field fill, rounded border and blue keyboard/controller focus cue as the other desktop controls. Their native bright borders and client edges are removed. API keys remain masked and editable when enabled, but the unintended vertical-scroll arrows are removed because single-line credentials have no scrolling action to expose.

**Modern-only desktop (batch 13)** retires the selectable classic desktop after the dialog modernization pass. Desktop startup, resizing, menus, Study windows, and owned dialogs now always use the modern presentation. Existing `modernLayout=0` preferences are migrated automatically, while models, prompts, profiles, capture settings, overlay settings, shortcuts, and other live values are preserved. Fullscreen remains a separate presentation and is unchanged. The original native controls and callbacks stay underneath the modern shell because they continue to provide the application behavior.

**Desktop control migration foundation (stage 1)** centralizes the nine desktop pages in one registry. Each page descriptor owns its identity, heading, sidebar destination, modern control group, adapted native controls, and layout callback. Navigation, visibility, and layout dispatch now consume that registry, while existing values and callbacks remain unchanged. This lets later stages replace adapted controls page by page without rewriting the shell or risking cross-page state.

**Game Text Translation control migration (stage 2)** creates the provider, model, prompt, capture, formatting, PNG-size, and startup controls directly inside the modern Game Text page. The page registry now owns that complete control tree and no hidden original form or duplicate model/prompt buttons are constructed. Shared control aliases remain available to the established translation logic and fullscreen presentation, so settings, capture actions, model/prompt management, and profile behavior stay synchronized while later stages continue the migration.

**Audio Translation control migration (stage 3)** creates the provider, model, output-language, audio-device, diagnostic, troubleshooting, and session controls directly inside the modern Audio page. The page registry owns the complete control tree; the hidden original Audio form, its legacy control group, and duplicate Add/Delete model buttons are no longer constructed. Existing aliases continue to feed audio startup, device testing, Profiles, and fullscreen controls, while model management is handled by the page-owned action menu.

**Explanation control migration (stage 4)** creates the provider, model, prompt, explanation, Study Library, saving, and Explainer-startup controls directly inside the modern Explanation page. The page registry owns the complete control tree; the hidden original Explanation form, its legacy control group, and duplicate model/prompt management buttons are no longer constructed. Existing aliases continue to feed explanation generation, Profiles, Study Library preferences, and fullscreen controls, while model and prompt management use the page-owned action menus.

**Overlay Windows control migration (stage 5)** creates opacity, color, typography, and move/resize controls directly inside the modern Translator and Explainer pages. Both registry descriptors own their complete control trees; the two hidden original overlay forms and their legacy control groups are no longer constructed. Shared aliases preserve Profiles, controller adjustment, live overlay theming, and independent Translator/Explainer settings, while the modern color buttons now serve as the shared color controls.

**Terminology control migration (stage 6)** creates the terminology enable switch, independent TL → TL and JP → TL profile selectors, and their manage/new/delete actions directly inside the modern Terminology page. Its registry descriptor owns the complete control tree; the hidden original Terminology form, duplicated help text, and legacy control group are no longer constructed. Existing aliases continue to feed screenshot translation, Profiles, fullscreen settings, and terminology editors while preserving glossary files and current selections.

**Profiles control migration (stage 7)** creates the saved-profile selector, startup-overlay selector, status text, and new/save/apply/delete actions directly inside the modern Profiles page. Its registry descriptor now owns the complete control tree; the hidden original Profiles form and duplicated explanatory text are no longer constructed. Existing aliases preserve profile file compatibility, active-profile selection, startup-overlay choices, and the same save/apply/delete behavior used by the fullscreen UI and external profile commands.

**Controls migration (stage 8)** creates the Keyboard/Controller view selectors, shortcut and controller-binding fields, row actions, conflict/status messages, and controller options directly inside the modern Controls page. Its registry descriptor owns both complete views and reveals only the selected one; the hidden original Controls form, duplicate row labels, and legacy responsive-layout scaffold are no longer constructed. Existing aliases preserve keyboard shortcut persistence, immediate conflict feedback, controller capture and assignment, D-pad navigation, fullscreen behavior, and the themed shortcut dialogs.

**API Keys migration (stage 9)** creates the in-app key toggle, masked Gemini/OpenAI fields, environment-variable action, save/delete actions, and About action directly inside the modern API Keys page. Its registry descriptor owns the complete functional control tree; the hidden original API Keys form, duplicate guidance, and legacy control group are no longer constructed. Existing aliases preserve portable `Settings/.env` compatibility, dirty/enabled states, Windows environment-variable guidance, fullscreen key management, and the shared About dialog.

**Advanced settings cleanup (stage 10)** removes the Paths page from both desktop and fullscreen navigation. `pythonExe`, `directModelOutput`, and `debugMode` remain supported in `Settings/control.ini`; direct model output and debug mode default to off. Existing legacy helper-path overrides continue to load for compatibility, but are no longer presented or rewritten by the interface.

**Desktop navigation paint and activation stability** keeps every modern button in the shared owner-draw renderer while keyboard arrows or controller D-pad focus moves through the interface. Focus changes now invalidate the existing themed surface instead of applying the retired native background/font effect, so traversed buttons retain their dark/light colors, rounded borders, typography, and sidebar icons. Enter, Space, and controller A/Cross invoke the focused owner-drawn button through its real click callback; Esc and controller B/Circle first cancel an active adjustment or open dropdown, then use the desktop window's existing close action. Keyboard/controller focus outlines remain visible, mouse focus behavior is unchanged, and fullscreen continues to use its separate renderer.

**Desktop section navigation order** makes Page Up/Down and controller shoulder navigation follow the modern sidebar: Game Text, Audio, Explanation, Translator overlay, Explainer overlay, Profiles, Terminology, Controls, and API keys. The separate Study Library window is intentionally not inserted into this page cycle, and forward/backward navigation wraps at the ends.

**720p Review for Anki layout** lowers the desktop Review window's minimum and initial height so native maximization can stay inside the Windows work area instead of being forced behind the taskbar. Below 760 pixels high, the title, filters, tabs, table, assessment reasoning, actions, and footer use a compact vertical arrangement. Assessment reasoning receives a taller scrollable area, while the candidate table keeps a practical minimum height.

Candidate selection is pointer- and controller-consistent: keyboard/controller selection changes use the row carried by the selection event, while a pointer click uses the ListView's exact clicked row. A single click on a sentence or vocabulary row therefore refreshes its AI assessment even if the native focused-row state has not caught up yet. Double-click remains the explicit shortcut for opening the selected entry in Reader.

The Sentences and Vocabulary segments use mutually exclusive selection styling. Pointer switching performs a settled repaint after the native click completes, preventing the previous segment from remaining visually active.

Study-only layout checks: `pwsh -NoProfile -File tests/test_desktop_layout.ps1 -StudyOnly`. They construct the real Study controls with temporary settings, verify compact/normal/wide layouts, button fit, screenshot bounds, editing controls and cleanup, and render synthetic screenshots without reading a personal library or contacting Anki. Use `-DialogsOnly` for the Anki, metadata, management, recommendation, shared message/menu, name-entry, model-selection, prompt, terminology, control-assignment, appearance, welcome, About, operational-message, and picker-polish checks. The full run also verifies the initial themed keyboard-capture display, in-app shortcut result messages, desktop action menus, focus-cue rendering, and shutdown with modern Study/Anki/authoring windows left open, including restoration of native table-header procedures and detachment of date-picker callbacks before GUI teardown.

Desktop page scrolling uses native mouse-wheel messages, coalesced into frames, rather than global wheel hotkeys. Rapid scrolling therefore does not consume AutoHotkey's hotkey flood limit; open dropdowns and opacity sliders retain their native wheel behavior. Page switches focus the selected section, and scrolling clips controls before moving them so text cannot briefly spill into the fixed header or footer.

On exit, pending scroll frames are cancelled and native wheel handlers are detached before GUI/global cleanup, preventing late destruction callbacks from opening an unassigned-variable error dialog. The desktop test suite includes an isolated live-window exit regression, normal parent/child destruction, and late callbacks with their tracking globals already released. Run only these checks with `pwsh -NoProfile -File tests/test_desktop_layout.ps1 -ShutdownOnly`.

The isolated desktop layout checks can be run with `pwsh -NoProfile -File tests/test_desktop_layout.ps1`; they exercise production layout functions against synthetic controls and temporary settings, without launching overlays or sending API requests. Audio coverage includes responsive fields, native dropdowns, tab order, inline diagnostic states, unclipped expanded help, session-state labels, and preservation of selections through page and modern relayouts. Explanation coverage adds independent AI choices, the original action callback, dependent Library/screenshot preferences, separate plain-text saving, and expandable startup settings. Overlay coverage checks both selectors, color values and targets, independent font/opacity settings, native size ranges, keyboard navigation, narrow layouts, move/resize callback wiring, and modern relayout persistence. Profiles/Settings coverage includes profile selection without applying, retained callbacks and values, both input views, duplicate-shortcut warnings, masked API keys and enabled states, advanced INI compatibility, preference migration, and the absence of classic-layout actions. Scroll regressions cover wheel-style steps and coalesced scrollbar dragging on all nine main pages, with opaque and translucent windows: the main window must never suspend its redraw state, and fixed header/sidebar/footer controls must not repaint. Modern desktop painting is buffered and invalidation is limited to the content viewport; fullscreen behavior remains unchanged.

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
`%TEMP%\JRPG_Overlay`. The live audio worker publishes its process marker under
`%TEMP%\JRPG_Control` so the desktop and full-screen controls can recover its
state after a launcher handoff or control-panel restart.

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
