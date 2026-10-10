# JRPG Translator LaunchBox / Big Box Integration

Version 1.0.0, for JRPG Translator 1.0.0.

This optional plugin starts JRPG Translator with your chosen settings when
you launch a game through LaunchBox or Big Box. It can also select a JoyToKey
mapping for that game. JoyToKey is optional; JRPG Translator also works without
the plugin.

[Open the illustrated setup guide](https://github.com/retrogamer0815/jrpg-translator-toolkit/blob/main/docs/manual/10-launchbox-and-big-box.md).

## Before you begin

Run **JRPG Translator.exe** directly and make a test translation first.
Configure your API key, models, capture region, overlays, and shortcuts.
Save your game-specific settings as a Profile if you want the plugin to load them.

Remember your **Show/Hide Control Panel** shortcut or controller binding:
the plugin normally starts JRPG Translator with its control panel hidden.

## Install

1. Close LaunchBox and Big Box.
2. Extract the plugin ZIP into LaunchBox's **Plugins** folder.
3. Check that the resulting path is:
   `LaunchBox\Plugins\JRPG Translator Integration\JrpgTranslator.LaunchBox.dll`.
4. Start LaunchBox.

## Set up a game

1. Right-click a game in LaunchBox and choose **JRPG Translator Setup...**.
   The same command is available in a game's details menu in Big Box.
2. Enable JRPG Translator and choose the Profile to load.
   **None — use current settings** keeps the application's current settings;
   it does not disable JRPG Translator.
3. Optionally enable JoyToKey and choose its profile.
4. If anything is missing, expand **Application locations** and browse to the
   relevant executable or JoyToKey profiles folder. Check the readiness
   indicators. **Ready** does not test your API key or internet connection.
5. Choose **Save**, then launch the game normally.

JRPG Translator Profiles and JoyToKey profiles are separate configurations;
their names do not need to match. Ensure JoyToKey's mappings match the keyboard
shortcuts currently configured in JRPG Translator.

## While playing

Use **Show/Hide Control Panel** to open the controls. The default keyboard
shortcut is **Ctrl+Shift+C**, unless you have changed it.

For a game launched through Big Box, this opens the fullscreen dashboard.
Use D-pad/arrows to move, A/Enter to confirm, and B/Esc to go back.
Use LB/RB or Page Up/Page Down to move through the full settings pages.
**Return to Game** takes you back to playing; **Open Desktop Interface**
opens the desktop controls.

## Startup and exit behavior

- If JRPG Translator is not running, the plugin starts it in the background
  and applies the chosen Profile before opening its configured startup overlays.
- If it is already running, the plugin applies the chosen Profile without
  restarting it. This does not automatically open or close its overlays;
  use their show/hide controls as needed.
- To change startup overlays for a Profile, use **Profiles > Startup overlays**
  in JRPG Translator, then **Save current**.
- If JoyToKey is enabled, the plugin starts it or switches an existing
  instance to the selected profile.
- When the game exits, the plugin closes only the program instances it
  started. Pre-existing JRPG Translator and overlay processes stay open.
  If JoyToKey was already running, its previous profile is restored where available.

## Moving the installation and backups

Locations inside the LaunchBox installation can be stored as relative paths
for portability. External locations may need updating after a move.
Reopen a game's setup to check the application locations and available profiles.

Back up the plugin's configuration with your LaunchBox installation, as well
as JRPG Translator's **Settings** folder. Translator Profiles alone do not
include the plugin's per-game selections. Windows environment-variable API keys
stay on the PC where you configured them and are not transferred with the folders.
