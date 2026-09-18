# 9. Profiles

[← Terminology Overrides](08-terminology-overrides.md) · [Contents](README.md) · [Next: LaunchBox and Big Box →](10-launchbox-and-big-box.md)

## What a Profile is for

A Profile remembers a set of game-related settings so you do not have to reposition overlays, reselect prompts, and rebuild capture choices every time you switch games.

It is a snapshot of selected settings, not a copy of the entire installation. Some settings deliberately remain global.

## Create and use a Profile

1. Configure the current capture, prompts, terminology, overlay appearance, and placement.
2. Select the Study Library you want to use for the game, if applicable.
3. Open **Profiles** and choose **New profile...**.
4. Give it a recognizable game name.
5. Check the selected Profile and use **Save current** when you want to store the current configuration.
6. To load an existing Profile from this page, select it and choose **Apply profile**.

> **Screenshot S16 — Profile management (to be added).** Show selection, Apply profile, Save current, New profile, and startup overlays.

### Selecting is not always applying

On the **Profiles page**, the dropdown selects the Profile to work with; **Apply profile** loads it. This allows you to choose a Profile before saving or deleting it.

In the **header Profile dropdown**, choosing a Profile is a quick apply action. Its final **Manage profiles...** entry opens the Profiles page.

This distinction prevents an accidental overwrite: confirm both the intended Profile name and the action before choosing **Save current**.

If switching prompts you about unsaved changes, choose deliberately between saving before switching, switching without saving, and cancelling. Current choices and saved Profile contents are not automatically identical.

## What is stored

| Stored in a game Profile | Kept as shared/global settings |
| --- | --- |
| Game Text and Explanation prompt selections | Provider/model selections for the AI functions |
| Capture mode, region, and target window information | Audio input device and output language |
| Guessed-subject and speaker-color toggles | Keyboard shortcuts |
| Selected glossary names and terminology enable setting | Direct controller action assignments/master enable |
| Translator/Explainer colors, fonts, sizes, bold, opacity, placement, and startup preferences | API keys and their storage configuration |
| Selected Study Library | Explanation-saving switches and capture-image-size limit |
| D-pad navigation setting | Other application-wide preferences |
| Associated Anki mapping/destination preferences when configured | Anki's actual decks, notes, templates, and database |

This table summarizes the current testing build. Do not rely on a Profile to preserve an entire old machine configuration.

In particular, **applying a Profile does not select a different AI provider/model for that game**. Check the shared AI settings if you changed them during another session.

## Startup overlays

Choose whether the next startup should open the Translator, Explainer, both, or neither, then save the Profile.

The LaunchBox plugin applies the selected Profile before opening startup overlays when it starts a new JRPG Translator instance. If the application is already running, applying a Profile does not automatically open or close its existing overlays.

This makes startup choices different from the footer's current-session visibility controls.

## Profiles and Study Libraries

A Profile can select the destination Library, while each saved explanation also records the active Profile as metadata. You might use:

- One Library for all games, with Profile filters separating them.
- A separate Library for each game.
- A Library for a course or study project containing material from several Profiles.

A Library is not created merely by giving a Profile a name. Manage Libraries in the Study Library window, then save the desired selection with the Profile.

## Delete, back up, or move Profiles

Use **Delete profile...** only for a Profile you no longer want. Deleting a settings Profile is not a command to delete its saved explanations, Anki cards, or game files.

Profile files are under `Settings\game_profiles`. Copying only one Profile file is not a complete migration: it can reference prompts, glossary profiles, Library names, installed fonts, and paths that are absent on another PC.

For a dependable backup, copy the complete Settings folder while the application is closed. If sharing a Profile publicly, review its paths and metadata and never include API credentials or private study material unintentionally.
