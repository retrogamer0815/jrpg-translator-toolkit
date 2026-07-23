using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using JrpgTranslator.LaunchBox;
using Unbroken.LaunchBox.Plugins;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string profileDirectory = Path.Combine(
            Path.GetTempPath(),
            "JrpgTranslatorLaunchBoxSmokeTest",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(profileDirectory);
        File.WriteAllText(Path.Combine(profileDirectory, "PC-Engine Tr.cfg"), string.Empty);

        string[] profiles = JoyToKeyProfileDiscovery.GetProfiles(profileDirectory).ToArray();
        Require(profiles.Contains("PC-Engine Tr", StringComparer.OrdinalIgnoreCase),
            "The expected JoyToKey profile was not discovered.");

        PluginConfiguration original = new PluginConfiguration();
        Require(string.Equals(
            original.TranslatorExecutable,
            @"Apps\JRPG Translator\JRPG Translator.exe",
            StringComparison.Ordinal),
            "The default JRPG Translator path is incorrect.");

        string translatorDirectory = Path.Combine(profileDirectory, "JRPG Translator");
        string translatorProfilesDirectory = Path.Combine(translatorDirectory, "Settings", "game_profiles");
        Directory.CreateDirectory(translatorProfilesDirectory);
        string translatorExecutable = Path.Combine(translatorDirectory, "JRPG Translator.exe");
        File.WriteAllText(translatorExecutable, string.Empty);
        File.WriteAllText(Path.Combine(translatorProfilesDirectory, "Retro Style.ini"), "[profile]\nschemaVersion=1\n");
        original.TranslatorExecutable = translatorExecutable;

        string[] translatorProfiles = TranslatorProfileDiscovery.GetProfiles(original).ToArray();
        Require(translatorProfiles.Contains("Retro Style", StringComparer.OrdinalIgnoreCase),
            "The expected JRPG Translator Profile was not discovered.");

        original.UpsertGame(new GameConfiguration
        {
            GameId = "smoke-test-game",
            GameTitle = "Smoke Test Game",
            TranslatorEnabled = true,
            TranslatorProfile = "Retro Style",
            JoyToKeyEnabled = true,
            JoyToKeyProfile = "PC-Engine Tr"
        });

        XmlSerializer serializer = new XmlSerializer(typeof(PluginConfiguration));
        using MemoryStream stream = new MemoryStream();
        serializer.Serialize(stream, original);
        stream.Position = 0;
        PluginConfiguration restored = (PluginConfiguration?)serializer.Deserialize(stream)
            ?? throw new InvalidOperationException("Configuration deserialization returned null.");

        GameConfiguration game = restored.GetGame("smoke-test-game", "Smoke Test Game");
        Require(game.TranslatorEnabled && game.JoyToKeyEnabled,
            "Per-game enabled settings did not round-trip.");
        Require(string.Equals(game.TranslatorProfile, "Retro Style", StringComparison.Ordinal),
            "The JRPG Translator Profile did not round-trip.");
        Require(string.Equals(game.JoyToKeyProfile, "PC-Engine Tr", StringComparison.Ordinal),
            "The JoyToKey profile did not round-trip.");

        restored.JoyToKeyProfilesDirectory = profileDirectory;
        GameSetupWindow setupWindow = new GameSetupWindow(restored, game.Clone());
        Require(string.Equals(setupWindow.Title, "JRPG Translator Setup", StringComparison.Ordinal),
            "The setup window did not initialize correctly.");

        IGameMenuItemPlugin menuItem = new GameSetupMenuItem();
        Require(menuItem.ShowInLaunchBox && menuItem.ShowInBigBox,
            "The game setup command is not enabled for both LaunchBox and Big Box.");
        Require(menuItem.IconImage != null,
            "The game setup command icon resource could not be loaded.");
        Require(!menuItem.SupportsMultipleGames,
            "The game setup command should only appear for a single selected game.");
        Require(string.Equals(menuItem.Caption, "JRPG Translator Setup...", StringComparison.Ordinal),
            "The game setup command has an unexpected caption.");

        IGameLaunchingPlugin runtimePlugin = new GameRuntimePlugin();
        MethodInfo[] runtimeMethods = runtimePlugin.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
        Require(runtimeMethods.Any(method => string.Equals(method.Name, "OnBeforeGameLaunching", StringComparison.Ordinal)),
            "The game launch lifecycle plugin was not initialized.");

        string testIniDirectory = Path.Combine(profileDirectory, "JoyToKeyState");
        Directory.CreateDirectory(testIniDirectory);
        string testIni = Path.Combine(testIniDirectory, "JoyToKey.ini");
        File.WriteAllText(testIni,
            "[General]\nFileName=Wrong Profile\n[LastStatus]\nFileName=Previous Profile\nOther=Value\n");
        string previousProfile = RuntimeProcessUtilities.ReadJoyToKeyActiveProfile(testIniDirectory);
        Require(string.Equals(previousProfile, "Previous Profile", StringComparison.Ordinal),
            "The previous JoyToKey profile was not read from LastStatus.");
        Require(RuntimeProcessUtilities.WriteJoyToKeyActiveProfile(testIniDirectory, "Restored Profile"),
            "The JoyToKey profile could not be restored in the test INI file.");
        string restoredProfile = RuntimeProcessUtilities.ReadJoyToKeyActiveProfile(testIniDirectory);
        Require(string.Equals(restoredProfile, "Restored Profile", StringComparison.Ordinal),
            "The restored JoyToKey profile did not persist in LastStatus.");
        File.Delete(testIni);
        Directory.Delete(profileDirectory, true);

        Console.WriteLine("Smoke test passed: {0} JRPG Translator Profile(s), {1} JoyToKey profile(s), configuration, setup window, menu command, and launch lifecycle verified.",
            translatorProfiles.Length,
            profiles.Length);
        return 0;
    }

    private static void Require(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
