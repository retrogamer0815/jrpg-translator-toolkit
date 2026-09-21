using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using System.Windows;
using System.Windows.Controls;
using JrpgTranslator.LaunchBox;
using Unbroken.LaunchBox.Plugins;

// Companion source snapshot: JRPG Translator v0.9.5.

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
        Require(setupWindow.Height <= SystemParameters.WorkArea.Height,
            "The setup window is taller than the visible desktop work area.");
        Require(setupWindow.MinHeight <= setupWindow.Height,
            "The setup window's minimum height prevents its adaptive startup size.");

        Grid setupRoot = (Grid)setupWindow.Content;
        ScrollViewer? setupScroller = setupRoot.Children
            .OfType<ScrollViewer>()
            .SingleOrDefault();
        Require(setupScroller != null
            && setupScroller.VerticalScrollBarVisibility == ScrollBarVisibility.Auto,
            "The setup window does not provide an automatic vertical scrolling region.");
        Require(!setupScroller!.CanContentScroll,
            "The setup window must use pixel-based scrolling so large sections remain reachable.");
        Require(Grid.GetRow(setupScroller!) == 1,
            "The setup window scrolling region is not in the expected flexible row.");

        StackPanel? setupFooter = setupRoot.Children
            .OfType<StackPanel>()
            .FirstOrDefault(panel =>
                Grid.GetRow(panel) == 2
                && panel.Children.OfType<Button>().Any(button =>
                    string.Equals(button.Content as string, "Save", StringComparison.Ordinal))
                && panel.Children.OfType<Button>().Any(button =>
                    string.Equals(button.Content as string, "Cancel", StringComparison.Ordinal)));
        Require(setupFooter != null,
            "The setup window Save/Cancel footer is not pinned outside the scrolling region.");

        CoverArtworkTests.Run(restored, game, profileDirectory,
            args.Length == 2 && args[0] == "--render-dir" ? args[1] : null);

        TestControllerDisclosureActivation(setupWindow);

        TestNoTranslatorProfile(setupWindow, restored, game, serializer, translatorProfilesDirectory);

        IGameMenuItemPlugin menuItem = new GameSetupMenuItem();
        Require(menuItem.ShowInLaunchBox && menuItem.ShowInBigBox,
            "The game setup command is not enabled for both LaunchBox and Big Box.");
        Require(menuItem.IconImage != null,
            "The game setup command icon resource could not be loaded.");
        using (Stream iconResource = typeof(GameSetupMenuItem).Assembly.GetManifestResourceStream(
            "JrpgTranslator.LaunchBox.Assets.menu-icon.png")!)
        {
            // Pin the supplied artwork so a stale embedded icon cannot pass a rebuild.
            string iconHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(iconResource));
            Require(iconHash == "68AB3C0F07BFFB251AA6C2CF9280A4DD9AD98D76AAE14F3D8F45ABFEF07779D5",
                "The menu must embed the updated JRPG Translator icon.");
        }
        Require(ReferenceEquals(menuItem.IconImage, new GameSetupMenuItem().IconImage),
            "The menu icon should be cached across menu instances.");
        Require(menuItem.IconImage!.Width == 1254 && menuItem.IconImage.Height == 1254,
            "The menu icon must retain the supplied resolution.");
        Require(!menuItem.SupportsMultipleGames,
            "The game setup command should only appear for a single selected game.");
        Require(string.Equals(menuItem.Caption, "JRPG Translator Setup...", StringComparison.Ordinal),
            "The game setup command has an unexpected caption.");

        IGameLaunchingPlugin runtimePlugin = new GameRuntimePlugin();
        MethodInfo[] runtimeMethods = runtimePlugin.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
        Require(runtimeMethods.Any(method => string.Equals(method.Name, "OnBeforeGameLaunching", StringComparison.Ordinal)),
            "The game launch lifecycle plugin was not initialized.");

        TranslatorGameContext dashboardGame = new TranslatorGameContext
        {
            Title = "The Legend of Xanadu",
            Platform = "NEC PC Engine-CD",
            BoxArtPath = @"C:\LaunchBox\Images\Xanadu.jpg",
            ClearLogoPath = @"C:\LaunchBox\Images\Xanadu Logo.png",
            PlatformLogoPath = @"C:\LaunchBox\Images\PC Engine.png",
            PlatformDevicePath = @"C:\LaunchBox\Images\PC Engine Device.png",
            PlatformDefaultArtPath = string.Empty
        };
        string[] bigBoxArguments = RuntimeProcessUtilities.BuildTranslatorArguments(
            translatorWasRunning: false,
            useBigBoxUi: true,
            translatorProfile: "Retro Style",
            gameContext: dashboardGame).ToArray();
        Require(bigBoxArguments.Contains("--background", StringComparer.Ordinal)
            && bigBoxArguments.Contains("--bigbox-ui", StringComparer.Ordinal)
            && !bigBoxArguments.Contains("--launchbox-ui", StringComparer.Ordinal)
            && !bigBoxArguments.Contains("--open-translator", StringComparer.Ordinal),
            "A cold Big Box launch did not receive the expected presentation-mode arguments.");
        int profileArgumentIndex = Array.IndexOf(bigBoxArguments, "--profile");
        Require(profileArgumentIndex >= 0
            && profileArgumentIndex + 1 < bigBoxArguments.Length
            && string.Equals(bigBoxArguments[profileArgumentIndex + 1], "Retro Style", StringComparison.Ordinal),
            "The JRPG Translator Profile was not preserved in the Big Box launch contract.");
        Require(ArgumentValue(bigBoxArguments, "--game-title") == dashboardGame.Title
            && ArgumentValue(bigBoxArguments, "--game-platform") == dashboardGame.Platform
            && ArgumentValue(bigBoxArguments, "--game-box-art") == dashboardGame.BoxArtPath
            && ArgumentValue(bigBoxArguments, "--game-clear-logo") == dashboardGame.ClearLogoPath
            && ArgumentValue(bigBoxArguments, "--platform-clear-logo") == dashboardGame.PlatformLogoPath
            && ArgumentValue(bigBoxArguments, "--platform-device-image") == dashboardGame.PlatformDevicePath
            && ArgumentValue(bigBoxArguments, "--platform-default-art") == string.Empty,
            "The running-game artwork context was not preserved in the Big Box launch contract.");

        string[] launchBoxArguments = RuntimeProcessUtilities.BuildTranslatorArguments(
            translatorWasRunning: true,
            useBigBoxUi: false,
            translatorProfile: string.Empty).ToArray();
        Require(launchBoxArguments.Contains("--launchbox-ui", StringComparer.Ordinal)
            && !launchBoxArguments.Contains("--bigbox-ui", StringComparer.Ordinal)
            && !launchBoxArguments.Contains("--open-translator", StringComparer.Ordinal),
            "A running JRPG Translator did not receive the expected LaunchBox presentation-mode reset.");

        foreach (bool wasRunning in new[] { false, true })
        foreach (bool bigBoxUi in new[] { false, true })
        foreach (string profile in new[] { string.Empty, "Retro Style" })
        {
            string[] startupArguments = RuntimeProcessUtilities.BuildTranslatorArguments(
                wasRunning, bigBoxUi, profile).ToArray();
            Require(!startupArguments.Any(argument => argument.StartsWith("--open-", StringComparison.Ordinal)),
                "The plugin must never override a Profile's startup overlay choices.");
            Require((ArgumentValue(startupArguments, "--profile") ?? string.Empty) == profile,
                "Startup must preserve the selected Profile or current-settings fallback.");
        }

        string[] clearContextArguments = RuntimeProcessUtilities
            .BuildTranslatorGameContextClearArguments()
            .ToArray();
        Require(clearContextArguments.Contains("--background", StringComparer.Ordinal)
            && clearContextArguments.Contains("--clear-game-context", StringComparer.Ordinal),
            "The running-game context cleanup request is incomplete.");

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

    private static void TestControllerDisclosureActivation(GameSetupWindow window)
    {
        Button locationsToggle = PrivateField<Button>(window, "_locationsToggle");
        StackPanel locationsPanel = PrivateField<StackPanel>(window, "_locationsPanel");
        Visibility initialVisibility = locationsPanel.Visibility;
        MethodInfo activateControl = typeof(GameSetupWindow).GetMethod(
            "ActivateControl",
            BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Controller activation helper was not found.");

        bool activated = (bool)(activateControl.Invoke(window, new object[] { locationsToggle }) ?? false);
        Require(activated && locationsPanel.Visibility != initialVisibility,
            "Controller activation must expand or collapse Application locations.");

        activated = (bool)(activateControl.Invoke(window, new object[] { locationsToggle }) ?? false);
        Require(activated && locationsPanel.Visibility == initialVisibility,
            "A second controller activation must restore the Application locations state.");
    }

    private static void TestNoTranslatorProfile(GameSetupWindow window,
        PluginConfiguration configuration, GameConfiguration originalGame,
        XmlSerializer serializer, string profilesDirectory)
    {
        ComboBox profile = PrivateField<ComboBox>(window, "_translatorProfile");
        Button refresh = PrivateField<Button>(window, "_refreshTranslatorProfiles");
        CheckBox enabled = PrivateField<CheckBox>(window, "_translatorEnabled");
        Require(profile.SelectedItem as string == "Retro Style",
            "An existing game Profile must remain selected when setup opens.");
        Require(profile.Items[0].ToString() == "None — use current settings"
            && profile.Items[0] is not string,
            "The first Profile choice must explicitly represent current settings, not a saved Profile name.");

        profile.SelectedIndex = 0;
        Require(window.Result.TranslatorProfile == "Retro Style",
            "Changing the dropdown must not save a Profile before the user selects Save.");
        refresh.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        enabled.IsChecked = false;
        enabled.IsChecked = true;
        Require(profile.SelectedIndex == 0,
            "Refresh and re-enabling Translator must preserve None instead of selecting a named Profile.");
        typeof(GameSetupWindow).GetMethod("SaveSelections", BindingFlags.Instance | BindingFlags.NonPublic)!
            .Invoke(window, null);
        Require(window.Result.TranslatorEnabled && window.Result.TranslatorProfile == string.Empty,
            "Saving None must clear the old Profile while leaving JRPG Translator enabled.");
        Require(window.Result.JoyToKeyEnabled
            && window.Result.JoyToKeyProfile == originalGame.JoyToKeyProfile,
            "None must not affect JoyToKey profile switching.");

        configuration.UpsertGame(window.Result);
        using MemoryStream saved = new MemoryStream();
        serializer.Serialize(saved, configuration);
        saved.Position = 0;
        PluginConfiguration loaded = (PluginConfiguration)serializer.Deserialize(saved)!;
        GameConfiguration game = loaded.GetGame(originalGame.GameId, originalGame.GameTitle);
        Require(game.TranslatorEnabled && game.TranslatorProfile == string.Empty,
            "The empty Profile must survive configuration save/reload.");
        GameSetupWindow reopened = new GameSetupWindow(loaded, game.Clone());
        Require(PrivateField<ComboBox>(reopened, "_translatorProfile").SelectedIndex == 0,
            "Reopening setup must show None for an empty saved Profile.");
        reopened.Close();

        foreach (bool running in new[] { false, true })
        foreach (bool bigBox in new[] { false, true })
        {
            string[] arguments = RuntimeProcessUtilities.BuildTranslatorArguments(
                running, bigBox, game.TranslatorProfile).ToArray();
            Require(!arguments.Contains("--profile", StringComparer.Ordinal)
                && !arguments.Contains("--open-translator", StringComparer.Ordinal),
                "None must launch with current settings, without a Profile or overlay override in either host.");
        }

        const string sameLabelProfile = "None — use current settings";
        File.WriteAllText(Path.Combine(profilesDirectory, sameLabelProfile + ".ini"), "[profile]\nschemaVersion=1\n");
        GameConfiguration collisionGame = game.Clone();
        collisionGame.TranslatorProfile = sameLabelProfile;
        GameSetupWindow collisionWindow = new GameSetupWindow(loaded, collisionGame);
        Require(PrivateField<ComboBox>(collisionWindow, "_translatorProfile").SelectedItem as string == sameLabelProfile,
            "A real Profile named like the None label must remain distinct and selectable.");
        collisionWindow.Close();

        PluginConfiguration noProfiles = new PluginConfiguration
        {
            TranslatorExecutable = Path.Combine(profilesDirectory, "empty", "JRPG Translator.exe")
        };
        GameSetupWindow emptyWindow = new GameSetupWindow(noProfiles, game.Clone());
        ComboBox emptyProfile = PrivateField<ComboBox>(emptyWindow, "_translatorProfile");
        Require(emptyProfile.Items.Count == 1 && emptyProfile.SelectedIndex == 0 && emptyProfile.IsEnabled,
            "None must be available even when no Profiles have been created.");
        emptyWindow.Close();
    }

    private static T PrivateField<T>(GameSetupWindow window, string name) where T : class
    {
        return (T)typeof(GameSetupWindow).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(window)!;
    }

    private static string? ArgumentValue(string[] arguments, string name)
    {
        int index = Array.IndexOf(arguments, name);
        return index >= 0 && index + 1 < arguments.Length
            ? arguments[index + 1]
            : null;
    }
}
