// Test-only LaunchBox/configuration boundary. Production runtime and ownership
// classes are compiled whole; this never reads real plugin settings.
using System;
using System.IO;
namespace Unbroken.LaunchBox.Plugins.Data
{
    public interface IGame { string Id { get; } string Title { get; } string Platform { get; } string FrontImagePath { get; } string ClearLogoImagePath { get; } string PlatformClearLogoImagePath { get; } }
    public interface IAdditionalApplication { }
    public interface IEmulator { }
    public interface IPlatform { string ClearLogoImagePath { get; } string DeviceImagePath { get; } string DefaultBoxImagePath { get; } string BackgroundImagePath { get; } }
}
namespace Unbroken.LaunchBox.Plugins
{
    using Data;
    public interface IGameLaunchingPlugin { void OnBeforeGameLaunching(IGame? game, IAdditionalApplication? app, IEmulator? emulator); void OnAfterGameLaunched(IGame? game, IAdditionalApplication? app, IEmulator? emulator); void OnGameExited(); }
    public interface ISystemEventsPlugin { void OnEventRaised(string eventType); }
    public static class SystemEventTypes { public const string LaunchBoxShutdownBeginning = "LB"; public const string BigBoxShutdownBeginning = "BB"; }
    public sealed class TestDataManager { public IPlatform? GetPlatformByName(string name) => null; }
    public static class PluginHelper { public static TestDataManager DataManager = new TestDataManager(); }
}
namespace JrpgTranslator.LaunchBox
{
    public sealed class PluginConfiguration
    {
        public string TranslatorExecutable = "";
        public string JoyToKeyExecutable = "";
        public string JoyToKeyProfilesDirectory = "";
        public GameConfiguration Game = new GameConfiguration();
        public GameConfiguration GetGame(string id, string title) => Game;
    }
    public sealed class GameConfiguration { public bool TranslatorEnabled; public bool JoyToKeyEnabled; public string TranslatorProfile = ""; public string JoyToKeyProfile = ""; }
    public static class ConfigurationStore { public static PluginConfiguration Value = new PluginConfiguration(); public static PluginConfiguration Load() => Value; }
    public static class PluginPaths
    {
        public static string DataDirectory = "";
        public static string ResolveTranslatorExecutable(PluginConfiguration configuration) => configuration.TranslatorExecutable;
    }
}
