# 12. Troubleshooting, backup, and advanced configuration

[← Anki](11b-anki.md) · [Contents](README.md)

## Start with a small test

When something stops working, separate capture/input, AI processing, and display:

1. Can the application capture the intended image or audio?
2. Does the selected provider/model accept the request?
3. Is the correct overlay visible and readable?

Change one setting at a time and preserve the exact error message. Repeated requests can consume API usage without helping identify the cause.

## Translation and API problems

| Symptom | First checks |
| --- | --- |
| No image translation | Confirm the capture target and visible dialogue; check key, provider/model, and overlay visibility |
| Wrong part of the screen translated | Reselect the region/window after resolution, monitor, window-position, or title changes |
| Authentication error | Check the selected provider's key, whitespace, project restrictions, and whether an older environment variable overrides the local key |
| Quota, billing, or rate-limit error | Check the provider's API project and usage dashboard; do not keep retrying indefinitely |
| Model unavailable or unsupported | Check the exact model identifier, task compatibility, and account access; refresh the model list |
| Names change between lines | Improve context and use appropriate Terminology Overrides |
| Japanese source missing for Explanation | Use a transcript-aware Game Text prompt and make a new translation |
| Odd speaker/subject formatting | Start from a bundled prompt with the expected markers/output structure |

A local “key present” or plugin “ready” indicator does not prove an API request will succeed. Conversely, a successful API response does not guarantee the result is displayed if the overlay is hidden.

Avoid sharing a complete environment file or unredacted error dump when asking for help.

## Audio problems

| Symptom | First checks |
| --- | --- |
| Local audio test detects no sound | Play voiced audio; choose the output device actually used by the game |
| Test succeeds but translation does not | Check the live-audio model, key, access, limits, and network |
| Audio stopped after connecting a headset | Stop translation, refresh devices, select/test the device, restart |
| Unrelated audio is translated | The selected playback device also receives another application's sound |
| Terminology change has no effect | Audio uses local target-language corrections; restart after glossary edits |
| Unexpected ongoing usage | Check whether the stream is still running even though the overlay is hidden |

Stop audio explicitly before closing only its output window or leaving the game idle.

## Windows, focus, and controllers

### Overlay hidden behind the game

Try borderless/windowed mode and check overlay visibility, opacity, and saved position. Exclusive fullscreen can prevent ordinary desktop windows from appearing over a game.

If you changed monitors, a Profile may restore old off-screen coordinates. Reposition the affected overlay and save the Profile again.

### Controller does not respond

Check **Settings → Controls** for detection and the relevant enable switch. Direct actions and D-pad navigation are separate. Verify that the intended app/dialog has focus for navigation.

If using a non-XInput device, check its driver/mapping configuration rather than assuming Xbox behavior. Test a keyboard shortcut to distinguish a controller-input issue from an action that is failing independently.

### One press moves twice or triggers two actions

Look for a native binding plus JoyToKey sending the same shortcut, duplicate D-pad-to-arrow mappings, or two active mapping tools. Remove the duplicate path and test again.

Remember that a direct controller action does not block the same input from reaching the game.

### Lower page controls seem unreachable

The desktop pages can extend below the first screen, especially at 720p. Continue through page content or scroll. If focus jumps past controls unexpectedly in a testing build, record the page, starting control, direction, window size, and Windows display scaling.

Report the navigation problem rather than saving a Profile as a supposed fix; Profiles do not define the desktop focus graph.

## Library and Anki problems

### Saved explanations seem missing

Check the selected Library, active search, filters, and date range. Clear filters and refresh. Check whether the material was saved under another Profile or Unsorted.

Confirm that **Save to Study Library** was enabled when the explanation was generated. A plain-text copy is not automatically a Library record.

### Screenshot missing

Screenshot saving may have been disabled, or its file may have moved. Enabling it now only affects future saves. Restore the relevant complete Library backup if you need a missing attachment.

### Recommendations are empty or “not assessed”

Check the review scope, filters, ignored vocabulary, and Anki exclusions. **New since last review** may exclude older material; try **All backlog**.

Opening the window does not generate AI recommendations automatically. Choose Generate and confirm the provider/model if you want an assessment.

### Anki connection fails

Start desktop Anki, confirm AnkiConnect is installed, and restart Anki after installation. Test the local connection again. Do not open public firewall/router ports as a workaround.

### Anki match missing or card blank

Check deck/subdeck scope, note type, and Japanese-field mapping. For a blank card, also check that the Anki card template displays the fields JRPG Translator fills.

Refresh link status after editing Anki. A manually set “Added to Anki” flag does not create or verify a note.

## Back up the data you care about

Close JRPG Translator and its study windows before copying active settings/databases. Stop the audio process too. Keep dated backups somewhere separate from the working installation.

The simplest comprehensive application backup is the **entire Settings folder**, plus any data configured outside it. It can include:

| Data | Usual location |
| --- | --- |
| Current application settings | `Settings\control.ini` |
| Saved game Profiles | `Settings\game_profiles` |
| Game Text prompts | `Settings\prompts` |
| Explanation prompts | `Settings\prompts_explain` |
| Terminology profiles | `Settings\glossaries` |
| Default/named/archived Libraries | `Settings\Study Library`, `Settings\Study Libraries`, `Settings\Study Libraries Archive` |
| Plain-text explanations | `Settings\Explanations` |
| Capture images | `Settings\Screenshots` by default; verify any custom location |
| Anki mapping/preferences | `Settings\anki.ini` |
| Shared candidate/ignored-vocabulary preferences | `Settings\anki_candidate_preferences.db` |
| In-app API credentials, if used | `Settings\.env` — contains secrets |

Also back up separately:

- Your Anki collection using Anki's own backup/export facilities.
- JoyToKey profiles.
- LaunchBox plugin/per-game configuration with the LaunchBox installation.
- Fonts or other external prerequisites as permitted by their licenses.
- Any custom paths outside the application folder.

A public “configuration backup” must not contain `.env`, private study data, or revealing absolute paths. A private backup containing credentials should be protected accordingly.

### Restore carefully

1. Keep a copy of the current state before replacing anything.
2. Close the app and stop processes that could write to the files.
3. Restore a coherent Settings/Library backup, not an arbitrary mixture of database files and attachments from different dates.
4. Open the matching application version where possible.
5. Verify the Library, Profiles, paths, and credentials before making new changes.

Removal operations may create local recovery backups and a Library Trash folder, but these live with the working data and are not a substitute for a separate backup. Do not prune them until you are sure you no longer need recovery.

## Advanced control.ini settings

The current UI intentionally has no Paths tab or show/hide-Paths setting. Most users should keep the supplied executable paths and diagnostics defaults.

Advanced users can configure these options in `Settings\control.ini`, under its existing `[cfg]` section:

```ini
[cfg]
pythonExe=.\python\python.exe
directModelOutput=0
debugMode=0
```

This is a **small example**, not a replacement for the whole file. Edit existing keys in the existing section; do not discard the rest of your settings or create conflicting duplicate sections.

### Python executable

`pythonExe` can point to the bundled interpreter or an already-installed environment's `python.exe`. Use an absolute path or a path relative to the application layout.

A custom interpreter must have the packages and compatible versions needed by the application's scripts. Pointing at an arbitrary system Python is not enough. The bundled runtime is the simplest baseline when diagnosing problems.

An external interpreter also reduces portability: moving the JRPG Translator folder does not move that external environment.

### Direct model output

`directModelOutput=1` bypasses the usual screenshot-translation output post-processing. It is intended for advanced testing or debugging prompt/model behavior.

It can change formatting and the structured output expected by other features. Do not enable it merely to improve appearance or as a general quality setting. The default is **0 (off)**.

### Debug mode

`debugMode=1` enables diagnostic logging. The control-panel log is normally under `%TEMP%\JRPG_Control\debug.log`; other components may produce their own diagnostics.

Logs may contain text, paths, or other sensitive session details. Review and redact them before sharing. Enable diagnostics for a specific reproduction, then turn them off again. The default is **0 (off)**.

### Safe editing procedure

1. Exit JRPG Translator so it cannot overwrite your edit on shutdown.
2. Back up `control.ini`.
3. Change only the intended key(s).
4. Save as a normal text INI file and restart.
5. Test one operation.
6. Revert from your backup if behavior is not as expected.

Do not alter unrelated path entries or delete settings as a first troubleshooting step.

## Report a useful issue

Use the [GitHub issue tracker](https://github.com/retrogamer0815/jrpg-translator-toolkit/issues). Include:

- Application version and whether it is a testing build.
- Windows version, screen resolution, and display scaling when relevant.
- Desktop or full-screen mode, plus LaunchBox/plugin involvement.
- Provider/model name, but **never the API key**.
- Exact steps, expected behavior, actual behavior, and the error text.
- A cropped/redacted screenshot or short recording if helpful.
- Whether the problem reproduces with one simple capture or a local audio test.

Attach only the minimum diagnostic information needed. Do not upload your entire Settings folder just to report a navigation bug.

[Back to manual contents](README.md)
