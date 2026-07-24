# JRPG Translator v0.9.1

Version 0.9.1 is a focused refinement release for the controller-first 0.9
workflow. It improves audio troubleshooting, removes an unnecessary screenshot
translation setting, and resolves several issues discovered while using the
tool alongside JoyToKey, LaunchBox, Big Box, and emulators.

## Highlights

### Test live audio before translating

The Audio Translation tab now includes **Test Audio**. Play any audible sound
and run the test to verify that the selected Windows output is reaching JRPG
Translator. The check is local and does not make an API request.

The tab also includes troubleshooting guidance for output-device selection and
emulator audio drivers. This is especially useful when one application can be
captured but an emulator using WASAPI or another backend cannot.

### Automatic screenshot output handling

The visible **Translation post-processing** menu has been removed. Prompts whose
names contain `with_transcript` or `with_kanji_reading` automatically use the
transcript-aware display path; other prompts use translation-only handling.

Advanced users can still bypass processing with the **Direct model output**
toggle in the optional Paths tab. It remains off by default.

### Cleaner controller and JoyToKey coexistence

The Controls tab now has a persistent **Use D-pad for control panel navigation**
option. Turn it off when another program maps the D-pad to keyboard arrows. This
prevents duplicate movement while keeping controller A / Cross and B / Circle
available. Profiles remember this choice, including Profiles selected through
the LaunchBox / Big Box plugin.

Esc now performs the same cancel or back action as controller B / Circle instead
of hiding the control panel.

### Prompt and terminology fixes

- Translation-only prompts no longer carry unnecessary transcript instructions.
- Multilingual explanation prompts have been synchronized.
- Repeated Japanese context in explanations includes hiragana readings for
  kanji.
- TL -> TL terminology overrides work again for screenshot translations.

### Additional polish

- Added opacity control for the entire control panel.
- Added dark-mode support to prompt, Profile, and hotkey entry dialogs.
- Documented prompt naming requirements in the New prompt dialog.
- Fixed online model picker cleanup errors and several capture, overlay,
  controller, caret, color, and refresh edge cases.

## Included source

The repository contains the AutoHotkey control panel and overlay, Python
translation scripts, bundled multilingual prompts, and LaunchBox / Big Box
plugin source. Personal API keys, Profiles, runtime settings, compiled
executables, backups, logs, screenshots, and generated user data are not
included.

For the complete change list, see `CHANGELOG.md`.
