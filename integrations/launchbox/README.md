# JRPG Translator LaunchBox / Big Box Integration

This preview plugin integrates JRPG Translator and JoyToKey with LaunchBox and
Big Box on a per-game basis.

Current behavior:

- Adds `JRPG Translator Setup...` to a game's LaunchBox context menu and Big Box details menu.
- Stores whether JRPG Translator should be used for that game.
- Detects JoyToKey `.cfg` profiles and stores an optional per-game profile.
- Lets users browse for the JRPG Translator executable, JoyToKey executable,
  and JoyToKey profiles folder from the per-game setup window.
- Stores LaunchBox-local paths relatively so a portable LaunchBox installation
  can be moved without editing the plugin configuration.
- Supports arrow, Enter, Space, and Escape navigation.
- Starts JRPG Translator in background mode before configured games.
- Starts JoyToKey or switches an existing instance to the configured profile.
- Closes only program instances started by the plugin after the game exits.
- Restores the previous JoyToKey profile when JoyToKey was already running.
- Leaves pre-existing JRPG Translator and overlay processes open.

## Install A Packaged Build

1. Close LaunchBox and Big Box.
2. Extract the plugin ZIP into `LaunchBox\Plugins`.
3. Confirm that the resulting path is
   `LaunchBox\Plugins\JRPG Translator Integration\JrpgTranslator.LaunchBox.dll`.
4. Start LaunchBox, right-click a game, and choose
   `JRPG Translator Setup...`.

The setup command is also available from a game's details menu in Big Box. Use
the browse buttons if JRPG Translator, JoyToKey, or the JoyToKey profile folder
is not in the automatically detected location.

## Build From Source

Building requires the .NET 9 SDK and a local LaunchBox installation. The script
checks `LAUNCHBOX_ROOT`, the common `%USERPROFILE%\LaunchBox` location, or an
explicit `-LaunchBoxRoot` argument.

Build with:

```powershell
.\build.ps1
```

The build output is under `bin\Release\net9.0-windows`. The build script also
runs a self-contained smoke test that does not use personal JoyToKey profiles.

Create an installable ZIP with:

```powershell
.\package.ps1 -Version 0.1.0-preview
```

Generated DLLs, ZIPs, logs, local plugin data, and machine-specific paths are
ignored by Git.

Do not place `manifest.json` beside the runtime DLL. LaunchBox reserves that
manifest flow for internally managed plugins; including it prevents ordinary
`IGameMenuItemPlugin` implementations from being registered in game menus.
The packaging script intentionally omits it.
