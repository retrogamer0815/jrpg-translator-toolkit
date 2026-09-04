# Big Box control center — Stage 5C2

Stage 3B replaces the prototype action grid with eight Home tiles and a shared
page-navigation framework. The correction explicitly separates frequent-use
Home shortcuts from the complete control-center pages. Stage 3C1 adds working
provider/model/prompt selectors for Translation AI and Explanation AI, and
provider/model/output-language selectors for Audio Translation AI. The same
selectors are available on the corresponding full L/R pages. Stage 3C2 adds
saving/startup preferences to the full Explanation page, without expanding the
Home Explanation AI shortcut. Stage 3C3 adds formatting/startup switches to the
full Screenshot Translation page; Home Translation AI also stays focused on
Provider, Model, and Prompt. Stage 3C4 adds device selection, device refresh, and
an inline audio test to the full Audio Translation page. Stage 3D enables direct
audio On/Off and a modern Capture view with region/window selection and PNG-size
adjustment. Stage 3E adds the complete Translation Window and Explanation Window
appearance pages and the Home overlay shortcut. Stage 3F replaces both Controls
previews with complete keyboard/controller binding pages and a controller-focused
Home shortcut. Stage 4 connects that shortcut to the existing Study Library and
adds reliable controller operation. Stage 5A gives that same Library a dedicated
fullscreen Big Box presentation while retaining its production data and actions.
Stage 5B applies the same presentation system to Study Reader while preserving
its version, section, editing, context, and Anki workflows.
Stage 5C1 ports Review for Anki and its complete recommendation setup flow.
Stage 5C2 adds a shared controller-first table mode to the fullscreen Library
and Review for Anki, a fullscreen Library column-layout editor, and direct
row actions for Anki candidates.
Advanced Settings remains available as the desktop fallback.

Stage 3G adds complete Terminology Overrides and Profiles pages with modern
controller-friendly editors. Stage 3H1 adds model-list management directly to
every Big Box Model picker. Stage 3H2 completes the remaining setup pages:
manual model entry, prompt-file management, API-key status/storage actions,
Paths, and a fullscreen About/help page. Keyboard input is still required for
secrets, paths, manual IDs, and substantial prompt editing, while every action,
confirmation, and return path remains controller accessible.

## Two independent navigation layers

Home retains eight shortcuts: Translation AI, Explanation AI, Audio Translation
AI, Audio Translation On/Off, Capture, Overlay Windows, Study Library, and Button
Configuration. Audio On/Off now acts immediately; the others open their own quick
views, not pages in the shoulder ring.
Quick views show a Home breadcrumb, hide the page arrows/indicator, and use Back
to return to their originating tile. Study Library opens as a separate managed
window and returns to Home when it closes. Capture is also reachable from the full Screenshot Translation page,
with Back returning to whichever entry point opened it.

The full L/R page order follows the desktop tabs:

Home → Screenshot Translation → Audio Translation → Translation Window →
Explanation → Explanation Window → Terminology Overrides → Profiles → Controls
→ API Keys → Paths.

Paths follows the existing optional `showPathsTab` setting. There are ten pages
including Home normally, or eleven with Paths enabled. Hiding Paths while it is
active safely returns to Home. Quick views and full pages remember focus
independently, even when they relate to the same feature.

During implementation, Home quick views and complete pages should reuse the same
settings operations and reusable controls. The long-term goal is complete
functionality in the Big Box presentation, making its Advanced Settings bridge
unnecessary. The normal desktop control center remains available independently.

## Controls

- LB/RB, L1/R1, or keyboard Page Up/Page Down: previous/next main page, wrapping
  at the ends. Holding a shoulder button changes only one page. These do not
  switch pages while inside a Home quick view; return to Home first.
- D-pad/arrows: spatial control navigation. The header's previous/next buttons
  are focusable and clickable. Tab/Shift+Tab traverse the active controls.
- A/Cross, Enter, or Space: activate the focused control.
- B/Circle or Escape: cancel an open choice view; otherwise return to Home;
  from Home, return to the game.
- Return to Game: leave directly from any page (cancel a choice view first).
- Focus is remembered per page, including the Home tile used to enter a page.

## Existing Study windows — Stage 4

The Home Study Library tile now hides the dashboard and opens the real Study
Library. Closing the Library returns to Big Box Home; desktop LaunchBox and the
standalone Study launchers retain their previous presentation. The existing
Study layout and all mouse/keyboard behavior are intentionally unchanged.

Native controller navigation covers the Study Library, Study Reader, Review for
Anki, and dialogs owned by those windows. D-pad movement follows the visible
control layout. Up/Down moves table rows or scrolls read-only explanation/source
text when those controls have focus. A/Cross activates buttons and opens the
selected Library or Anki-review row. B/Circle closes a dropdown first, then the
current dialog or Study surface through its normal close path.

In the Study Reader, LB/RB browses the previous/next saved explanation. In
Review for Anki, LB/RB switches between Sentences and Vocabulary and puts focus
in the corresponding table. Shoulders are consumed in the main Library because
it has no peer tabs. JoyToKey-mirrored arrows, Enter/Space, Escape, and Page
Up/Down are suppressed only for the matching physical controller action while
a Study surface is active; genuine keyboard input continues to work normally.

Initial focus is placed in the Library table, Reader explanation, or active Anki
candidate table so the most common controller action works immediately. Focus
cues use the native control repaint only, avoiding a full-window redraw on every
D-pad press.

## Fullscreen Study Library — Stage 5A

Opening Study Library from Big Box now presents the production Library in a
borderless fullscreen shell on the game monitor. It uses the dashboard's header,
spacing, colors, controller legend, active-library display, and a dedicated
**Back to Dashboard** action. The table, filters, search, context pane, Library
management, export, Anki, chapter, and Reader actions are the existing controls
and implementations rather than a second copy of the Library logic.

At normal and large fullscreen sizes, the screenshot and Original Japanese
viewers share the lower context area side by side. This gives 4:3 captures a
substantially larger aspect-ratio-aware preview while retaining the source text.
Compact resolutions fall back to a taller stacked preview. Table columns scale
to physical fullscreen pixels and the Japanese source receives spare width;
these presentation widths are not written over the desktop column settings.

Desktop LaunchBox and both standalone Study launchers continue to use the
resizable desktop Library. Big Box fullscreen coordinates are never written to
the saved desktop bounds, and the desktop first-run welcome is not placed over
the controller presentation. The initial fullscreen frame is DWM-cloaked until
its layout is complete, avoiding a visible desktop-sized intermediate window.

## Fullscreen Study Reader — Stage 5B

Opening an entry from the fullscreen Library now opens Study Reader in a matching
borderless fullscreen shell on the same monitor. The header identifies the active
Library and provides a visible **Back to Library** action. The footer documents
Reader-specific controls, including LB/RB entry browsing. Closing with the button
or B/Escape restores the Library, activates it, and returns controller focus to
the table.

The fullscreen shell reuses the production Reader controls and handlers: entry,
version and explanation-section navigation; the manual Anki marker; new version,
copy, Add to Anki and editing workflows; screenshot browsing; and original source
text. The explanation is the primary large surface. Context uses a side-by-side
screenshot/source layout so a 4:3 capture remains useful and undistorted without
making the Japanese source unavailable.

The desktop and standalone Reader retain their resizable window, title bar and
saved position. Fullscreen bounds never overwrite those preferences. If a Reader
is already open in a different presentation, it is safely recreated rather than
trying to add or remove a shell from live controls. As with the Library, the first
fullscreen frame remains DWM-cloaked until layout, data, theme and image work are
complete, avoiding a desktop-sized intermediate flash.

## Fullscreen Review for Anki — Stage 5C1

Opening Review for Anki from the fullscreen Library now keeps the user inside the
same borderless Big Box presentation on the game monitor. The production sentence
and vocabulary tables, Show and AI filters, refresh, review completion, vocabulary
triage, Anki actions, context menu and recommendation cache remain shared with the
desktop window. The fullscreen layout gives flexible table width to Japanese
sentence text or vocabulary meaning/context and scales the remaining columns for
the active display.

The active Library appears in the header and **Back to Library** is a permanent,
visible action. A centered **Sentences**/**Vocabulary** page heading, `1 / 2`
indicator, underline and visible ‹/› buttons make the second table discoverable;
LB/RB performs the same switch. Opening an entry in the fullscreen Reader and
closing it now returns to Review for Anki and restores focus to the active candidate
table instead of skipping back to Library.

The batch recommendation confirmation, learner-level/selection-style choices,
advanced study-focus preferences, optional guidance editor and complete prompt
preview all receive matching fullscreen Big Box shells. They continue to use the
existing settings and bridge commands. Generating recommendations retains the
looping indeterminate progress bar, candidate-count message and disabled
**Generating, please wait…** button because the provider does not expose reliable
per-item completion. Single-entry generate/regenerate and recommendation deletion
remain available through the controller-visible table context menu.

Desktop Review for Anki and all its recommendation dialogs keep their resizable or
compact desktop presentation. Fullscreen windows are composed while DWM-cloaked,
so none of these transitions intentionally exposes a small intermediate window.

## Shared Study table mode — Stage 5C2

The fullscreen Study Library and Review for Anki now both expose **Table mode...**.
It replaces the surrounding toolbar and context pane with one large table while
it is active. Column widths use comfortable presentation minimums instead of
shrinking every field into the viewport, so long Japanese sources, grammar and
vocabulary context remain readable. The ‹/› buttons and D-pad Left/Right scroll
the table horizontally; Up/Down continues to browse rows. Review for Anki keeps
LB/RB available for switching between its Sentences and Vocabulary tables, and
the table-mode heading identifies which one is active.

A opens the selected Library explanation. In fullscreen Review for Anki, A on a
focused sentence or vocabulary row instead opens **Row actions**, with **Add to
Anki...** first, followed by Open in Reader, recommendation, review, and
ignore/restore commands as applicable. This preserves the focused table row, so
adding vocabulary no longer requires moving focus away from its selection.
**Selected actions...** exposes the same menu from the table-mode action row;
both paths reuse the existing commands rather than maintaining duplicate Anki or
recommendation implementations. Desktop Review retains its direct-open behavior.
The row counter and disabled Open action clearly cover empty tables. B/Escape or
**Exit table mode** restores the normal fullscreen layout, focus, footer, prior
horizontal scroll position, and Library screenshot preview.

The fullscreen Library **Columns...** action now opens a controller-oriented
column-layout editor. The overview identifies every column's current order,
visibility, and width. Selecting one opens large actions for moving it earlier or
later, changing width in 20-pixel steps, showing/hiding optional columns, and
resetting its width. **Reset all defaults** is also staged in the editor. All
changes remain a working copy until **Save layout**; Cancel leaves the persisted
desktop/Big Box layout untouched. Saving writes through the existing Library
column settings, so order, visibility, and width stay shared with the desktop
Library rather than becoming presentation-only preferences.

The resizable desktop Library and Review layouts do not expose or inherit the
Big Box-only editor or table-mode presentation state.

Page transitions reuse the existing fullscreen window. A 150 ms indicator
animation does not slide, hide, fade, or change the opacity of that window.
The timer stops on completion, hiding, or shutdown. Disabled/owned modal dialogs
do not route shoulder input to the underlying dashboard. In-page choice views
also block the main page navigation until the choice is applied or cancelled.

## AI selectors — Stage 3C1

- Choose Provider, Model, or Prompt/Output language to open a fullscreen-styled
  in-page choice view. Up to four choices are shown as tiles. With five or more
  choices, the view becomes a vertical scrolling list with five reusable rows;
  there are no Previous/Next choices buttons. The current value is marked and
  initially focused.
- The long-list focus lens keeps the pending value in the center, with two
  progressively smaller neighboring rows on each side. The outer rows also use
  muted text; the center retains the app's native themed focus button. This is
  a bounded-control presentation, not a 3D rendering layer or a new animation
  timer. The item counter shows the precise position in the list.
- Up/Down or the mouse wheel scroll one value at a time, stopping at each end.
  Page Up/Down scroll five items; Home/End reach the first/last item. Left/Right
  or Tab reach Cancel, and Up/Down from Cancel restore the pending item. Shoulder
  buttons still do not change full control-center pages while choosing.
- A/Enter selects the center item; clicking any visible row selects that exact
  row. Mouse-wheel capture is limited to an active long-list view while the
  pointer is over the dashboard, so it cannot intercept scrolling elsewhere.
- Browsing, cancelling, hiding the dashboard, and selecting the current value
  do not change settings. A/Enter/Space or a mouse click explicitly applies a
  different choice. Held controller buttons do not confirm a second action.
- Cancelling or confirming returns focus to the exact tile that opened the
  picker. Its return target is stored with the picker, independently of general
  page focus memory. Closing temporarily suppresses focus-memory updates while
  the parent controls are restored, so Windows focusing Back to Home during
  hide/disable cannot replace that target. The same behavior covers device lists.
- Choices are read from the actual desktop dropdowns. Applying a value uses
  the existing `AutoPersist`, provider-enablement, `ApplyShotSettings`, or
  independent `ExplainPromptChanged` routines, not a second settings file.
- Each provider retains its own model. Translation and explanation prompts
  remain independent. Audio changes apply on the next start; changing settings
  does not interrupt or restart live audio.
- Profile/provider changes and removed list entries are rechecked before
  applying. Stale choices are rejected with an inline message. Empty lists and
  save failures also produce inline feedback instead of a blank selector or
  false success message.
- Model-list management and prompt editing remain in Advanced Settings.
  Audio-input controls are on the full Audio Translation page. Capture/PNG-size
  controls are shared between Home and the full Screenshot Translation page.

## Explanation preferences — Stage 3C2

The full Explanation page combines Provider, Model, and Prompt with five switches:
Save to Study Library, Source screenshots, Save plain-text copies, Open Explainer
on startup, and Open always on top. The Home Explanation AI shortcut still shows
only its three AI selectors. Each switch shows On/Off; focusing it displays a
short explanation, and a successful change produces an inline saved notice.

The full page now has three labeled rows with subtle separators: **AI settings**
(Provider, Model, Prompt), **Study Library** (Library, Screenshots, Plain-text copies),
and **Explainer startup** (Open on startup, Always on top). Navigation follows
these groups. Group captions are beside the rows so a small screen does not lose
height to extra headings; the grouped buttons use slightly smaller text with
room for both their label and current value. AI pickers and Home quick views
retain their existing layout and font sizes.

Both presentations use the desktop checkboxes and a shared single-key INI setter.
The Big Box setter writes before changing the displayed value, so a failed write
retains the old state and shows an inline error. No new database or overlay actions
are introduced. Disabling Library saving does not remove existing content; it
disables the screenshot option while preserving its previous preference. Disabled
options are skipped by controller navigation, including when desktop/profile
changes disable the currently focused option.

The two startup switches retain the desktop behavior: opening on startup applies
the next time JRPG Translator starts, and always-on-top applies when the Explainer
next opens. Toggling them does not launch or modify an already open overlay. Model
management, prompt editing, explanation generation, and study actions remain
available through Advanced Settings during migration.

## Screenshot Translation preferences — Stage 3C3

The full Screenshot Translation page now uses the same three-row presentation:

- **AI settings:** Provider, Model, Prompt.
- **Translation formatting:** Highlight guessed subjects, Use speaker name color.
- **Startup:** Open Translator on startup, Open always on top, Clear screenshots
  on startup.

The five switches use their existing desktop checkboxes and INI keys through a
shared single-key setter. Writes complete before Big Box changes its displayed
value. Formatting switches also call the existing `ApplyShotSettings`, updating
the environment for the next screenshot translation; no new request is started
and no existing text is regenerated. Desktop checkbox changes and profile-applied
values are reflected in the dashboard.

Startup switches only persist preferences. They do not immediately launch an
overlay, change the topmost state of an open overlay, or delete files. Cleanup's
contextual help explicitly explains that the existing next-start routine deletes
captures recorded from the previous session. Merely opening either presentation
does not change any preference. Home Translation AI retains only its three AI
selectors, and Screenshot-only switches cannot be activated from Explanation,
Home quick views, an AI picker, or while a modal/write guard is active.

Capture selection and maximum PNG size are implemented together in Stage 3D.
The screenshot action buttons and model/prompt management still use Advanced
Settings until their respective modern workflows are ready. The Explanation
section caption is now **STUDY LIBRARY**, as requested; plain-text copies remain
independent of Library saving despite sharing that organizational row.

## Audio Translation input — Stage 3C4

The full Audio Translation page has three groups:

- **AI settings:** Provider, Model, Output language.
- **Audio input:** Listen device, Refresh devices, Test audio.
- **Status:** non-interactive, inline progress and diagnostic results.

Listen device uses the same native device list as the desktop panel and the
existing in-page selector (tiles for up to four choices, scrolling for longer
lists). An explicit selection is saved before its displayed value changes. It
applies on the next audio start, without restarting or redirecting a live session.
Home's Audio Translation AI shortcut remains limited to its three AI selectors.

Device refresh and audio testing now share a hidden asynchronous diagnostic
worker in both presentations. Competing input controls are disabled while the
worker runs, with a busy label and an inline status. A 30-second deadline prevents
the UI waiting indefinitely for a stalled driver. The process handle is retained
for targeted timeout/shutdown cleanup; no unrelated process is killed by looking
up a possibly reused PID. A unique temporary result file is read only after the
helper exits and removed after completion. Closing the app stops its diagnostic.
Navigating away from the Audio page does not abandon an active check.

Refresh preserves a disconnected selection and reports it as unavailable instead
of silently switching inputs. Errors keep the previous device list. Tests report
detected sound, silence, device errors, missing results, or timeout without a
blocking popup. Test audio makes no AI request and refuses to run during live
Audio Translation; it never stops that session. Device refresh is still allowed.
The status text is not a controller focus target.

This step updates `scripts/live_audio_translator.py` alongside the executable.
Its optional device-list result file preserves the existing stdout format used
by startup discovery. This is not the broader asynchronous-Python migration:
initial device population and translation execution are otherwise unchanged.

## Audio On/Off and Capture — Stage 3D

The Home audio tile starts/stops immediately without opening a preview page.
The full Audio Translation page has the same action in its Audio session row.
Busy feedback and a shared guard prevent duplicate presses; each held A press
acts once. Input diagnostics must finish before toggling audio. The existing
status timer updates On/Off and external state changes without redrawing an
unchanged button. Actual model/service errors remain in the Translator output.
The busy audio tile stays enabled to retain native focus and its blue outline;
the shared busy guard, not disabling the focused button, blocks repeat actions.
Successful startup shows the same non-activating corner notification as the
desktop panel, alongside the inline Big Box status. Stop retains its existing
corner notification. Failed or already-running starts do not announce a new On.

Both presentations share audio startup. Big Box gets inline missing-path,
missing-key, launch-error, and early-exit messages. It does not use the desktop
startup recovery scan or synchronous retry after a failed start. The native Run
call now explicitly captures its output PID, and recovery scans exclude audio
diagnostics as well as device enumeration. Stop still uses the existing stop
routine; this is not a general asynchronous translation-engine migration.

Capture is a modern in-page view with Select region, Select window, and Maximum
PNG size. It shows the current target and opens no desktop mode-selection dialog.
The full Screenshot Translation page exposes the same view alongside formatting
options. Its breadcrumb and Back destination remember whether it came from Home
or Screenshot Translation.

Region/window selection reuses the existing overlay picker and controller/mouse
behavior. A preset/saved region is immediately adjustable without needing an
initial mouse drag. The dashboard hides during selection and returns on confirm
or cancel, preserving the game return target and priming controller edge state.
The desktop capture watcher is unchanged. Big Box watches only the explicit
completion sequence, so unrelated settings writes cannot finish a selection.
The overlay writes completion status before that sequence marker. Sending has
a 2.5-second response timeout; missing completion has a 125-second return limit,
slightly beyond the overlay's existing 120-second auto-cancel. Shutdown stops the
dashboard watcher without reopening a GUI. This workflow selects a target; it
does not take a screenshot or make an AI request.

Maximum PNG size uses a separate in-page adjustment view with ±100/±500 KB
buttons, a large pending value, and Save/Cancel. Values are bounded to 100–10,000
KB. Back, Escape, or hiding the dashboard cancels pending changes. Save writes
only the existing capture/maxKB key, then synchronizes the runtime value and
desktop field. An unchanged value causes no write. Failed writes, profile changes,
or externally changed size settings cannot silently overwrite the previous value.

## Overlay appearance — Stage 3E

The full Translation Window and Explanation Window pages mirror their desktop
tabs in three groups: Colors, Text, and Window. Translator settings include
window/text/speaker colors, font, font size (6–128), bold, opacity, and move/resize.
Explainer has the same controls except speaker color, with its existing font-size
range of 6–200. These pages reuse the desktop font lists and the same runtime
variables, native fields and INI sections (`cfg` and `cfg_explainer`). The Home
Overlay Windows shortcut first selects Translator or Explainer, then exposes
window color, opacity and move/resize; it remains outside the shoulder-page ring.
Its two chooser tiles distinguish Closed, Visible, and Hidden. Hidden includes
both a genuinely hidden native window and the overlay's safe sent-behind state,
where the live window remains technically visible but is no longer topmost.

Color, opacity and font size have in-page Save/Cancel editors. Color uses the
existing hue/saturation/brightness conversion and native gradient drawing, plus
a local color swatch. Opacity and font size show their pending numeric value.
Up/Down navigates; Left/Right adjusts the focused slider. Holding Left/Right
repeats, accelerating for colors and opacity but not font size. A/Enter advances
from a slider; the explicit Save button applies the setting. Mouse dragging also
works. Native trackbars have no AHK Focus event, so a guarded, one-shot mouse
callback updates the shared focus outline after Windows processes a slider click.

Previewing or cancelling never changes the real overlay or settings. Save first
checks that the settings path, profile, persisted key and current runtime value
have not changed. It writes only that setting, synchronizes its native desktop
control, and sends the existing theme payload only to the relevant overlay. This
new send has a 2.5-second bound; failure to respond is reported as saved-but-not-
applied, rather than leaving the dashboard waiting indefinitely. An unchanged
value does not write or send. Font selection uses the existing virtual list for
more than four choices and preserves the opening tile on confirm or cancel.
Bold is an immediate, guarded toggle. Existing desktop handlers are unchanged.

Move/resize calls the existing native controller adjustment workflow with an
explicit Big Box return context. The dashboard hides while the overlay is moved;
Save keeps the new bounds, Cancel restores the old bounds, and both return to the
same tile. Held confirmation is primed on return. Missing overlays fail inline,
partial setup failures clean up the adjustment, and quiet shutdown does not
reopen either control center. Desktop callers keep their existing return path.

## Controls and Button Configuration — Stage 3F

The full Controls page now has three groups: **Configure** opens the complete
Keyboard shortcuts or Controller bindings list; **Controller** contains the
global Direct action bindings and D-pad navigation switches; and **Status** shows
the existing controller status without adding a focus stop. The Home Button
Configuration shortcut deliberately omits keyboard shortcuts, retaining only
the controller list, the two controller-wide switches, and status.

Both binding lists contain the same ten actions and friendly labels as Advanced
Settings. They reuse the existing five-row scrolling picker rather than creating
ten permanent tiles. Selecting an action opens a focused detail page showing its
current binding. Keyboard actions provide Change shortcut, Disable, and Restore
default; controller actions provide Assign controller button and Disable. Back
returns first to the same action in the list, then to the tile that opened it.

Keyboard entry uses a native Hotkey field inside the fullscreen dashboard with
explicit Save and Cancel. Modifier order is canonicalized only for comparison,
so equivalent shortcuts such as Ctrl+Shift+E and Shift+Ctrl+E cannot bypass the
duplicate check. Applying a shortcut writes the existing `[hotkeys]` key,
updates the desktop field, signals the overlay reload file, and calls the same
four live rebind routines already used by Advanced Settings.

Controller assignment remains inside the fullscreen design. It first waits for
all buttons to be released, then displays the detected button and saves when it
is released. A short B/Circle press may itself be assigned; holding B/Circle for
one second cancels. Esc, the Cancel button, controller disconnect status, and a
60-second timeout provide other safe exits. The normal action/navigation poll is
suspended only while this capture owns input and is reset afterward.

Duplicate keyboard shortcuts or controller buttons open a modern conflict page.
Nothing changes until **Move binding** is confirmed; **Keep current bindings**
and B/Esc are non-destructive. Move validates both persisted values before
clearing the old action and writing the new one, with rollback if the second
write fails. All individual writes similarly compare the opened value with the
current runtime, desktop control, and INI value, preventing a stale dashboard
from overwriting external changes. Hiding the dashboard cancels pending capture
or editing and restores the parent page on the next open.

## Model management — Stage 3H1

Every Model picker now has a **Manage models…** action. It opens a nested Big
Box-styled view with **Add from catalogue…**, **Add model ID manually…**,
**Refresh online catalogue…**, and **Remove a local model…**. The catalogue uses
the provider and purpose belonging
to the opening Translation, Explanation, or Audio page, so all six underlying
model lists remain independent. Existing IDs are filtered out before selection.
Left/Right reaches the management action from anywhere in a long Model list;
pressing Down once more at its final item reaches it as well.

Catalogue loading runs as a cancellable background child process. The title
animates while it is active, timeout and provider errors remain in the dashboard,
and online, fresh-cache, and older-cache results are identified. Only the exact
catalogue child and its unique temporary result are cancelled or cleaned up.
Adding a model persists it through the existing list format, selects it, and
updates the matching desktop/runtime setting. Removing a model requires a second
confirmation and never permits the final local model to be removed. Each write
validates the active provider, profile, and original model-list snapshot first.
Manual model entry validates empty and duplicate IDs, persists through the same
provider/purpose list, and selects the new ID after a successful save.

## Setup and maintenance — Stage 3H2

Every Translation and Explanation Prompt picker now has **Manage prompts…**.
The nested fullscreen workflow creates, edits, backs up, and deletes prompt files
without opening the desktop control center. New prompt files are not created
until their text is explicitly saved, and deletion always requires confirmation.
At least one prompt is retained. Substantial prompt text entry intentionally uses
a native multiline keyboard field; all surrounding selection and confirmation
remains controller friendly.

The API Keys page shows only provider status and storage source—never a key value.
Each provider tile can write or replace only its own entry in `Settings/.env`,
preserving unrelated lines and making an atomic backup first. Entry is masked and
nothing is written until Save. Provider-specific removal and deletion of the
whole in-app key file require confirmation. Windows user environment variables
are opened through the existing system action and are never changed or deleted
by the dashboard. The About view provides the version, guide/video, project and
bug-report links, plus the existing copyable diagnostic version information.

The optional Paths page exposes all five existing executable/script fields plus
Direct model output and Debug mode. File paths use an explicit editor with Browse
and Save; a path that does not currently exist requires a second **Save anyway**
confirmation, matching the desktop behavior. The two switches persist
immediately. Every successful update synchronizes the existing desktop controls
and runtime values instead of introducing a second configuration source.

Setup notices are scoped to their page, secrets never enter labels/notices, and
pending values are discarded on Back, dashboard hide, or shutdown. Nested views
return focus to the exact Model/Prompt/API/Path action that opened them.

## Dashboard focus and repainting

All dashboard buttons now use a thick accent outline in addition to the native
focus cue: bright blue in dark mode and dark blue in light mode. Four disabled,
non-focusable strips sit outside the focused button without covering its text or
intercepting button clicks. They scale with the layout, follow controller,
keyboard, and native focus changes, and disappear when the dashboard is hidden.
There is no focus-animation timer, opacity change, or native button replacement.

Previously, every D-pad focus step called the generic dialog helper, which forced
`RedrawWindow` with `RDW_ALLCHILDREN | RDW_UPDATENOW` on the entire dashboard.
The Big Box path now leaves normal old/new button invalidation to Windows and
reveals native focus cues only when the UI-state flags require it. Repeated Focus
notifications do not rewrite unchanged help text, and stale notifications cannot
move the highlight back to an earlier control. Other desktop dialog focus logic
is deliberately unchanged. See Microsoft's documentation for
[RedrawWindow](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-redrawwindow)
and [WM_CHANGEUISTATE](https://learn.microsoft.com/en-us/windows/win32/menurc/wm-changeuistate).

## Automated checks (Windows, AutoHotkey v2)

Run `tests/test_bigbox_navigation.ps1`. Optional parameters are `-AutoHotkey` and
`-OutputDirectory`. It first loads the complete script with an early ExitApp for
syntax validation, then extracts the production dashboard and controller
navigation code into an isolated harness with hidden native controls. Production
settings persistence and screenshot-environment updates run against synthetic
controls and an INI file in the test output directory. External application
services are stubbed: the test does not read personal configuration, call models,
launch overlays, poll a real controller, or modify a Study Library.

Checks compare the complete page list against the actual desktop tab labels and
order. They also cover optional Paths visibility, all eight independent quick
views, page-versus-quick-view input routing, focus restoration, modal
guards, XInput/legacy shoulder navigation, held A/B/shoulder behavior, no preview
side effects, timer cleanup, title ellipsis, artwork aspect ratios, and layout
bounds/overlaps at 1280x720, 1920x1080, and 3840x2160. Generated PNGs render only
the synthetic hidden test GUI. Native app theming is stubbed, so their button
colors are not a visual dark-mode acceptance test.

AI checks additionally cover the six page/quick-view entry points, list scrolling,
current-choice no-ops, independent model/prompt settings, desktop synchronization,
real INI writes, runtime environment updates, cancellation, held confirm, stale
choices, empty lists, and a deliberately unwritable settings destination. Long-list
checks exercise the 4/5-choice cutoff, a 100-item prompt list, exact item mapping,
clamped boundaries, focus/cancel traversal, controller and wheel input, row bounds
and font heights, held A/B, and both XInput and legacy controller confirmation.

Explanation preference checks cover all five single-key writes, desktop/Big Box
synchronization, retained screenshot preferences, navigation around a disabled
option, held controller A, quick-view/chooser/modal isolation, failed writes, and
the absence of overlay or study side effects. The grouped-layout/focus follow-up
suite, extended through Stage 5B, picker-focus restoration, and audio feedback,
has 3,926 Big Box assertions plus 79 focused Study-controller assertions
and also covers the Audio page.
A native child-window paint probe first verifies that
it detects a deliberately forced full repaint, then checks that normal focus
changes do not repaint unrelated Home, Explanation, Screenshot Translation, or Audio tiles. This uses a cloaked,
off-screen test window without activating it; it does not take over the user's
foreground app. Additional checks cover group navigation, non-interactive outline
geometry, native/stale focus callbacks, light/dark colors, font restoration after
leaving Explanation, help text height, and outline cleanup/restoration on hiding.

Screenshot-specific checks cover all five persistence mappings, immediate runtime
formatting flags, independence from Explainer/Library settings, held controller A,
desktop/profile synchronization, quick-view/page/chooser/modal isolation, and
unwritable destinations retaining the old checkbox/runtime state. Group captions
and two-line buttons are checked for clipping at all three resolutions.

Audio checks cover Unicode device selection, real single-key INI writes, desktop
synchronization, cancellation, failed writes, missing devices, list deduplication,
quick-view isolation, and refusing to test during live audio. A small AHK child
fixture exercises the production hidden-process/timer code without accessing an
audio driver: success, silence, error, missing result, held confirmation, timeout,
shutdown cleanup, process-handle cleanup, temporary-file cleanup, and completion
after leaving the page. Inherited environment variables are restored immediately
after launch. Result text is checked for clipping at all three resolutions.

A native hide/disable probe also injects the transient Back to Home focus that
could previously overwrite a picker's return target. Tests cover all three AI
fields in the three Home quick views and their full pages, plus the audio-device
picker: XInput/legacy cancellation, held B, Escape, unchanged confirmation, and
changed confirmation. They verify native focus, the accent outline, page memory,
no writes when cancelling, and rejection of stale Focus callbacks. This regression
test failed against the old close sequence and passes with the protected target.

Stage 3D checks cover direct/held audio actions, inline errors, modal/diagnostic
guards, external state updates, capture entry/return paths, nested PNG cancellation,
size bounds, persistence, stale/failed writes, missing helpers, failed sends,
explicit capture completion, cancellation, same-target reselection, ignored
unrelated writes, timeout, and shutdown cleanup. The actual capture suspend
routine hides only the synthetic test GUI; overlay communication/resuming the
foreground dashboard are stubbed. A separate short-lived AHK child exercises the
real audio-start routine, verifying its PID and absence of scans or retries.
Audio feedback regressions observe native focus and the blue outline while the
action is still in flight, not only after completion, on Home and the full Audio
page in both themes. Start, stop, failed start, and held A are covered. The real
startup routine is also checked for exactly one On notification after success
and no new On notification for invalid configuration, an early exit, or an
already-running helper. These tests reproduced both reported issues before the
respective fixes.
Neither test calls Python audio drivers or an AI service. Capture and PNG editor
layout are rendered at 720p, 1080p and 4K; unchanged periodic status is also checked
with the native paint probe.

Stage 3E checks cover all appearance mappings and both overlay sections, desktop
control synchronization, bounded/targeted theme requests, Save/Cancel and no-write
previews, stale profiles/settings, failed writes, exact original colors, font
choice lists, numeric bounds, held buttons, mouse/controller slider focus and
cleanup. The production gradient renderer is used for visual checks at 720p,
1080p and 4K. The production move/resize start/finish routines operate only on a
synthetic native window, verifying cancellation, confirmation, return routing
and shutdown. Actual game focus restoration and overlay delivery remain manual
checks; the harness never opens a real game or contacts an AI service. Native
paint probes confirm that navigating either appearance page leaves unrelated
tiles unpainted.

Stage 3F checks cover full and Home Controls layouts at 720p, 1080p and 4K;
controller-wide toggles; the ten-action scrolling lists; exact INI/runtime/desktop
synchronization; assign, disable and restore-default paths; held-B capture;
disconnect and timeout feedback; duplicate cancellation and atomic moves; stale
and failed writes; live keyboard reloads; nested focus restoration; and dashboard
hide/reopen cleanup. Controller snapshots and settings are synthetic and confined
to the temporary harness directory.

Stage 3H1 checks all six provider/purpose/list mappings, bounded Model-picker
navigation, catalogue filtering and source notices, immediate selection after
adding, active-model fallback after removal, last-model protection, confirmation
and nested Back paths, stale snapshots, provider errors, older-cache results,
forced refresh, loading cancellation, and dashboard-shutdown cleanup. Catalogue
responses are synthetic; these checks make no network or AI request.

Stage 3H2 checks manual model IDs; grouped API and Paths layouts; masked key entry
and removal without exposing secrets; provider-specific `.env` preservation;
valid, missing, stale and failed path saves; immediate diagnostic/output toggles;
prompt creation, editing with backup, deletion, cancellation and last-prompt
protection; About actions; and nested controller return paths. All files and
configuration live in the temporary harness directory. Browser/system actions,
file pickers, API services, and personal configuration are stubbed.

Run `tests/test_audio_input_protocol.py` with Python for seven additional protocol
tests. They extract only the diagnostic functions, using fake devices rather than
importing audio/AI dependencies or reading local configuration. They cover Unicode,
deduplication, stdout compatibility, empty lists, driver errors, newline safety,
result-file write errors, and the existing audio-test result format.

## Manual acceptance checks

1. Launch a game from Big Box and open the dashboard. Confirm game artwork,
   title, profile, and Home tile summaries are correct.
2. Visit every main page using each shoulder, including wraparound. Check the
   small indicator animation and verify that the desktop never flashes through.
3. Open a tile with A, return with B, and verify focus returns to that tile.
   Hold A or B during a transition; there must be no second activation.
4. Reach the header arrows with the D-pad and test mouse clicks as well.
   Check Page Up/Down, Tab/Shift+Tab, Enter/Space, and Escape on a keyboard.
5. Compare Home's Translation AI tile with the full Screenshot Translation page
   reached by R from Home: the former has a Home breadcrumb and no page arrows;
   the latter has full-page navigation and a page count. Repeat for the other
   shortcuts. B returns to the correct Home tile. Advanced Settings opens the
   desktop panel intentionally.
6. Reopen the dashboard, check focus memory and game focus restoration, and
   repeat at the available monitor resolutions and Windows scaling settings.
7. Launch through desktop LaunchBox and confirm it still opens the normal
   desktop control center with its existing controller tab navigation.
8. From Home, try all three AI shortcuts. Change provider, model, and prompt or
   output language. Check the matching full L/R page and Advanced Settings show
   the same values. Restart the app and confirm the selections were saved.
9. Open a model, prompt, or language list with more than four choices. Scroll
   directly with Up/Down and the mouse wheel; test the first/last entries and
   click a neighboring row. Check the font-size effect and item counter. Cancel
   with B/Esc and confirm nothing changed and focus returns to the original
   setting. Hold A while confirming and B while cancelling: each press should
   perform only one action. A list with four or fewer items must still use tiles.
10. Generate a translation and explanation from the game with the new settings.
    For audio, stop/start audio translation to use new provider/model/language
    selections. Verify the current session is not interrupted merely by choosing.
11. Reach the full Explanation page with L/R. Toggle each saving/startup option
    and compare Advanced Settings. Reopen the app to check persistence. Confirm
    the Home Explanation AI shortcut still contains only Provider, Model, Prompt.
12. Disable Save to Study Library: Source screenshots should be disabled and
    skipped by D-pad navigation. Re-enable saving: its previous screenshot choice
    should return. Existing Library entries and screenshots must remain intact.
13. Hold A on a switch: it must toggle only once per press. Check contextual help
    and saved feedback. The two startup switches must not immediately open or
    alter an overlay; verify their behavior on the next applicable start/open.
14. Check the three Explanation groups at your normal resolution and Windows
    scale. Move quickly through Home tiles, Explanation rows, page arrows, and
    choice lists with the D-pad. The blue outline should clearly identify focus;
    unrelated tiles should not flash. Repeat with Tab and mouse clicks, and check
    light mode. Returning from a picker or hiding/reopening must retain the right
    focus without a stale outline.
15. On the full Screenshot Translation page, change the two formatting options
    and compare Advanced Settings. Generate a new translation to check their
    effect. Hold A: each physical press should toggle only once. Verify Home
    Translation AI still has just its three AI selectors.
16. Check the three startup preferences match the desktop and survive restarting.
    Toggling alone must not open/reorder an overlay or delete captures. Only enable
    screenshot cleanup if you want its existing next-start deletion behavior;
    there is no need to enable it just to test navigation. Check the Explanation
    section now reads STUDY LIBRARY and its preferences remain unchanged.
17. On the full Audio Translation page, choose a listening device. Check that
    Advanced Settings shows the same selection, cancelling a picker changes
    nothing, and restarting retains your explicit choice. A running audio session
    must not restart just because its next-start device was changed.
18. Refresh devices and check its busy label and inline completion message. If
    practical, disconnect the selected output and refresh: the selection should
    remain, with an unavailable notice. Reconnect/refresh or choose another device.
19. Stop Audio Translation, play game audio, and use Test audio. Check the inline
    result, then repeat with no sound. While Audio Translation is running, Test
    audio should explain that it must be stopped first without stopping it for
    you. No AI call should be made by a test. Navigate away during a check and
    return; controls should recover when it finishes.
20. From Home, start and stop audio with its tile. Hold A briefly: one press must
    cause only one action. Try the same control on the full Audio page and check
    the displayed state follows your usual audio hotkey too. The outline should
    remain on the audio tile throughout, without flashing on the left arrow.
    Both successful On and Off should show the usual upper-left notification.
21. Open Capture from Home. Select a region, move/resize it with your controller,
    and confirm. The dashboard should return to that capture tile. Repeat and
    cancel with B; the previous target should remain. Repeat with Select window.
    Check the mouse works as well, and test a normal translation afterward.
22. Open Capture from the full Screenshot Translation page. Back should return
    there, not Home. Adjust Maximum PNG size and cancel; then adjust and Save.
    Compare the desktop value and restart the app to confirm persistence. Check
    that holding the confirmation/cancel button does not activate another action
    as the dashboard returns from the overlay selector.
23. Open both full Window pages with L/R. Change window/text colors, Translator
    speaker color, font, font size, bold and opacity. Compare the corresponding
    desktop controls. Only the chosen overlay should change. Restart to verify
    persistence. An overlay that was closed should use the saved setting when
    next opened; choosing an appearance setting should not launch it.
24. Open a color editor and adjust all three gradients with the D-pad and mouse.
    Cancel with B, then repeat with Save. Try opacity and font size similarly.
    Pending values must not affect the actual overlay before Save. Cancel should
    preserve the exact original value and return to the tile that opened it.
    Check that font selection also returns to Font when cancelled or confirmed.
25. From Home, open Overlay Windows and choose each overlay. Check the three
    quick actions and the nested Back path. Confirm that L/R still traverses the
    full settings pages only, not these quick views or pending editors.
26. Test move/resize from both the full pages and quick views, including Save
    and Cancel. The dashboard should hide during movement and return to the same
    tile, without opening the desktop panel. Hold A briefly when confirming;
    returning must not immediately start another adjustment. Verify the original
    desktop Move / Resize buttons still return to the desktop panel.
27. Move focus between appearance tiles in dark and light mode, including mouse
    clicks on sliders. The blue outline should follow correctly and unrelated
    tiles should not flash. Check the layout at the available display scales.
28. Open the full Controls page with L/R. Verify Keyboard shortcuts, Controller
    bindings, both controller switches, and status are present. Open Button
    Configuration from Home and confirm it contains the controller controls only.
29. Browse all ten controller actions, assign an unused button, and compare the
    Controls tab in Advanced Settings. Disable it and confirm both presentations
    update. Hold B during assignment to cancel; briefly press and release B to
    confirm it can still be assigned intentionally.
30. Assign a controller button already used by another action. Keep the current
    bindings first, then repeat and choose Move binding. Confirm no duplicate is
    left and both the old and new action update. Restart to verify persistence.
31. In Keyboard shortcuts, change, disable, and restore one shortcut. Confirm the
    desktop Controls tab and live shortcut follow immediately. Try an existing
    shortcut and verify the conflict page changes nothing until Move is confirmed.
    Hide/reopen the dashboard during controller capture and check it returns to
    Button Configuration without a stale capture screen.
32. Open Terminology Overrides with L/R. Toggle overrides and select separate
    TL → TL and JP → TL profiles; compare every value with Advanced Settings and
    restart to verify persistence. Open and cancel a profile picker to confirm
    focus returns to the tile that opened it.
33. Add, edit and delete terminology entries from the fullscreen manager. Check
    that deletion requires confirmation, cancelling preserves the entry, and the
    default profile cannot be deleted. Create/delete a non-default profile and
    confirm the other glossary type remains intact. If practical, place a malformed
    line in a test glossary and verify the raw repair editor preserves it on Cancel.
34. Open Profiles with L/R. Select a profile and verify it is not applied until
    Apply is chosen. Save current settings, create a safely named profile, and test
    both overwrite and delete confirmations. Compare the desktop profile list and
    verify a failed or cancelled action leaves the original profile intact.
35. In every Stage 3G text editor, use A / Enter while the field is focused; it
    should move to the next action rather than inserting an unexpected command.
    B / Circle / Esc must cancel pending text or confirmation without saving.
36. Open Model from each Translation, Explanation, and Audio quick/full page.
    Select Manage models and verify the four controller actions appear without
    opening a desktop window. B should return first to Model, then to its page.
37. Add a model from the catalogue. During a slow request, check that the title
    keeps moving and Cancel remains usable. Confirm the source/cache notice is
    clear, existing models are excluded, and the added model is selected in both
    Big Box and Advanced Settings. Restart once to verify persistence.
38. Remove a non-current model, cancel at confirmation, then remove it for real.
    Repeat with the active model and verify a remaining model is selected. A list
    containing one model must refuse removal. Test an unavailable network/API key
    and confirm the error and retry stay inside the Big Box design.
39. From every Model picker, open Manage models and add a valid model ID manually.
    Confirm duplicates and blank IDs are rejected, the new ID is selected in the
    matching desktop list, and B returns to the Model tile without changing it.
40. From Translation and Explanation Prompt, open Manage prompts. Create a test
    prompt, edit it, and remove it after testing both cancellation paths. Confirm
    Advanced Settings sees the same files and that the last prompt cannot be
    deleted. Use B from each nested screen and verify focus returns correctly.
41. Open API Keys with L/R. Confirm provider tiles reveal only Configured/Missing
    and Windows/In-app—not key contents. Test a temporary in-app key, cancel first,
    then Save and remove it. Verify Windows variables and unrelated `.env` lines
    remain unchanged. Check the About links and copied version information.
42. If Paths is enabled, edit one path and cancel, then Browse/Save a valid file.
    Try a deliberately missing path and cancel at **Save anyway**. Toggle Direct
    model output and Debug mode, compare Advanced Settings, then restore both.
    Ensure a notice from API Keys never appears on Paths or another main page.
43. From Big Box Home, open Study Library. Confirm there is no small-window flash,
    the Library fills the game monitor with the modern header and controller
    footer, and the dashboard hides. Switch libraries and confirm the active
    Library value in the header updates. Verify the table has focus, Up/Down
    moves rows, and A opens the selected entry in Study Reader. At the first/last
    real row, press Up/Down again and confirm
    focus leaves the table for the nearest control instead of becoming trapped.
    Move Left/Right across both toolbar rows and confirm focus remains in the same
    visual row; use Up/Down to change rows. At a horizontal edge, focus should
    remain on the edge control rather than jumping diagonally or wrapping. Close the Library with B
    and verify Big Box returns to Home with one press—without briefly restoring
    Advanced Settings or the game. Repeat with **Back to Dashboard**, and at
    1280×720, 1920×1080, and 3840×2160 where available; controls must remain
    readable without overlap. Finally open the Library outside Big Box and
    confirm its normal resizable layout and previously saved bounds are intact.
    In both presentations, open the Library selector and move with Up/Down:
    browsing must not switch libraries until A/Enter confirms the highlighted
    choice; B must close it without changing the active Library.
    Check a 4:3 source screenshot is large and undistorted, the table occupies
    its width instead of showing compressed columns, and controller focus on
    Version or Original Japanese does not highlight all of their text.
44. Open Study Reader from the fullscreen Library. Confirm it fills the same game
    monitor without a small-window flash, uses the matching modern header/panel/
    footer, identifies the active Library, and keeps the explanation and context
    readable without overlap. A 4:3 screenshot must remain undistorted beside the
    original Japanese. Use Up/Down to scroll a long explanation and LB/RB to browse
    entries. At the first/last scroll position, press Up/Down once more and verify
    focus returns to the nearest control; short Original Japanese fields which do
    not need scrolling must never trap focus. Open its Anki/copy/edit menus and confirm D-pad visibly changes the
    highlighted menu row, A selects it, and B closes it. In the Anki setup dialog,
    open a native dropdown and confirm Up/Down plus A selects an entry. Repeat in
    other owned dialogs. Open Anki connection/link check and immediately close it
    while the initial check is still running; its late result must be discarded
    without an AutoHotkey error or another dialog. Closing Reader with B or its
    Back to Library button should reveal and focus the fullscreen Library again.
    Repeat outside Big Box and confirm the desktop Reader remains resizable and
    returns to its previously saved bounds.
45. Open Review for Anki from the fullscreen Library. Confirm it fills the game
    monitor without a desktop-sized flash, identifies the active Library, and its
    table columns remain readable at 1280×720, 1920×1080 and 3840×2160 where
    available. Confirm that the heading, `1 / 2` indicator, underline and ‹/›
    buttons make both Sentences and Vocabulary visible and switchable. Use LB/RB
    for the same switch and Up/Down to choose rows. Press A on a sentence and a
    vocabulary row: a visibly highlighted row-action menu must open with Add to
    Anki first, and choosing it must use the row that remained selected. Confirm
    Open in Reader and the applicable recommendation/ignore actions are present,
    then close with B or Back to Library. Open a candidate in Reader,
    close Reader, and verify focus returns to the same Review table rather than
    Library. Test the context menu's single-entry generate/regenerate and delete
    recommendation actions. Start batch recommendations and confirm its setup,
    Customize and prompt-preview screens all retain the fullscreen design and
    controller navigation. During generation, verify the looping progress bar,
    candidate-count message and Generating button remain visible until completion.
    If JoyToKey also maps these buttons, each physical press must move or activate
    only once; ordinary keyboard and mouse input must remain unchanged. Repeat
    outside Big Box and confirm the original desktop layouts are unchanged.
46. In fullscreen Library, enter Table mode. Confirm the context pane and normal
    toolbars are replaced by one large table, its title and row counter are clear,
    Up/Down browses rows, and Left/Right plus ‹/› reveal the deliberately wider
    columns. A should open the selected explanation; Selected actions should open
    the normal Library context menu. Exit with B and verify the original layout,
    screenshot, focus and horizontal position return. Open Columns in the
    fullscreen Library. Move one column, change its width, hide an optional
    column, and cancel; confirm nothing changed. Repeat and Save layout; confirm
    order, width, and visibility update in the Library and persist after reopening
    it. Required Japanese source must not be hideable, and Reset all must remain
    staged until Save. Repeat Table mode in Review for Anki for both Sentences and
    Vocabulary, switching them with LB/RB. Verify A opens row actions without
    losing the chosen candidate, Add to Anki works for vocabulary, and B exits
    table mode before a second B returns to Library. Finally confirm neither
    desktop Study window shows the Table mode action and desktop Review still
    opens a candidate directly with A/Enter.

## Updated roadmap

| Stage | Scope |
| --- | --- |
| 3B correction | Separate Home quick views from the complete L/R settings pages. |
| 3C1 (implemented) | Shared working AI selectors in three Home quick views and their full pages. |
| 3C2 (implemented) | Saving/startup switches on the full Explanation page; Home Explanation AI stays focused. |
| 3C3 (implemented) | Screenshot formatting/startup switches in grouped rows; Home Translation AI stays focused. |
| 3C4 (implemented) | Audio input selection, asynchronous device refresh, and bounded inline audio diagnostics. |
| 3C remaining | Any remaining full-page actions, mapped to their matching workflows below. |
| 3D (implemented) | Direct audio toggle and shared window/region capture selection, including maximum PNG size. |
| 3E (implemented) | Complete Translation Window and Explanation Window pages, plus quick overlay adjustments. |
| 3F (implemented) | Complete Controls page for keyboard/controller settings, plus the quick controller-binding view. |
| 3G (implemented) | Complete Terminology Overrides and Profiles pages, including modern profile/entry editors, inline failures, stale-file checks and destructive-action confirmations. |
| 3H1 (implemented) | Controller-friendly model management: add from the online catalogue and remove from local model lists, with a dedicated Big Box-styled view, asynchronous loading, cache/error status, stale-state checks and safe confirmation. |
| 3H2 (implemented) | API Keys, Paths, manual model entry, prompt-file management and About/help actions in controller-accessible fullscreen workflows. |
| 4 (implemented) | Reliable controller operation in the existing Study Library, Reader, Review for Anki, and their owned dialogs. |
| 5A (implemented) | Modern fullscreen Big Box Study Library presentation using the existing Library data and actions, while preserving the desktop window. |
| 5B (implemented) | Modern fullscreen Big Box Study Reader presentation using the existing Reader workflows, with responsive context, explicit Library return, no first-frame flash, and protected desktop bounds. |
| 5C1 (implemented) | Fullscreen Review for Anki, responsive sentence/vocabulary tables, Reader return routing, and fullscreen batch-recommendation confirmation, customization and prompt preview. |
| 5C2 (implemented) | Shared controller-first table mode, fullscreen Library order/width/visibility editor, and direct Review row actions with Add to Anki first. |
| 5C3 | Remaining Big Box Study transitions and modern owned-dialog workflows outside recommendation setup. |
| 5D | Multi-resolution Study presentation testing and polish. |
| 6 | Full functionality audit and integration testing; retire the Big Box Advanced Settings bridge only once nothing depends on it. |
