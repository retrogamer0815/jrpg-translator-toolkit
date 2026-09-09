using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Unbroken.LaunchBox.Plugins;
using Unbroken.LaunchBox.Plugins.Data;

namespace JrpgTranslator.LaunchBox
{
    public sealed class GameRuntimePlugin : IGameLaunchingPlugin
    {
        public void OnBeforeGameLaunching(IGame? game, IAdditionalApplication? app, IEmulator? emulator)
        {
            // LaunchBox invokes these callbacks for additional applications too. The
            // integration belongs to the main game launch and must only run once.
            if (game == null || app != null)
            {
                return;
            }

            try
            {
                RuntimeCoordinator.Start(game);
            }
            catch (Exception exception)
            {
                // A plugin failure must never prevent LaunchBox from starting the game.
                RuntimeLog.Write("Unhandled game-launch callback failure: " + exception.Message);
            }
        }

        public void OnAfterGameLaunched(IGame? game, IAdditionalApplication? app, IEmulator? emulator)
        {
        }

        public void OnGameExited()
        {
            try
            {
                RuntimeCoordinator.StopActiveSession();
            }
            catch (Exception exception)
            {
                // Cleanup is best-effort and must never destabilize the host process.
                RuntimeLog.Write("Unhandled game-exit callback failure: " + exception.Message);
            }
        }
    }

    public sealed class RuntimeShutdownPlugin : ISystemEventsPlugin
    {
        public void OnEventRaised(string eventType)
        {
            if (string.Equals(eventType, SystemEventTypes.LaunchBoxShutdownBeginning, StringComparison.Ordinal)
                || string.Equals(eventType, SystemEventTypes.BigBoxShutdownBeginning, StringComparison.Ordinal))
            {
                try
                {
                    RuntimeCoordinator.StopActiveSession();
                }
                catch (Exception exception)
                {
                    RuntimeLog.Write("Unhandled host-shutdown callback failure: " + exception.Message);
                }
            }
        }
    }

    internal static class PluginHostEnvironment
    {
        public static bool IsBigBoxHost()
        {
            try
            {
                return string.Equals(
                    Process.GetCurrentProcess().ProcessName,
                    "BigBox",
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
    }

    internal static class RuntimeCoordinator
    {
        private static readonly object Sync = new object();
        private static RuntimeSession? _activeSession;

        public static void Start(IGame gameInfo)
        {
            lock (Sync)
            {
                string gameId = gameInfo.Id ?? string.Empty;
                string gameTitle = gameInfo.Title ?? string.Empty;
                if (_activeSession != null)
                {
                    if (string.Equals(_activeSession.GameId, gameId, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    RuntimeLog.Write("A new game launch replaced an unfinished runtime session.");
                    StopSession(_activeSession);
                    _activeSession = null;
                }

                try
                {
                    PluginConfiguration configuration = ConfigurationStore.Load();
                    GameConfiguration game = configuration.GetGame(gameId, gameTitle);
                    TranslatorGameContext gameContext = RuntimeProcessUtilities.CreateGameContext(gameInfo);
                    bool useBigBoxUi = PluginHostEnvironment.IsBigBoxHost();
                    if (!game.TranslatorEnabled && !game.JoyToKeyEnabled)
                    {
                        return;
                    }

                    RuntimeSession session = new RuntimeSession(gameId, gameTitle, gameContext);
                    _activeSession = session;

                    if (game.TranslatorEnabled)
                    {
                        StartTranslator(configuration, game, session, useBigBoxUi);
                    }

                    if (game.JoyToKeyEnabled)
                    {
                        StartJoyToKey(configuration, game, session);
                    }
                }
                catch (Exception exception)
                {
                    RuntimeLog.Write("Game launch preparation failed: " + exception.Message);
                }
            }
        }

        public static void StopActiveSession()
        {
            RuntimeSession? session;
            lock (Sync)
            {
                if (_activeSession == null)
                {
                    return;
                }

                session = _activeSession;
                _activeSession = null;
            }

            // Do potentially slow process cleanup after releasing the coordinator
            // lock. A repeated exit/shutdown callback will now see no active session
            // and cannot clean up the same processes twice.
            StopSession(session);
        }

        private static void StartTranslator(
            PluginConfiguration configuration,
            GameConfiguration game,
            RuntimeSession session,
            bool useBigBoxUi)
        {
            string executable = PluginPaths.ResolveTranslatorExecutable(configuration);
            if (!File.Exists(executable))
            {
                RuntimeLog.Write("JRPG Translator executable was not found: " + executable);
                return;
            }

            session.TranslatorExecutable = executable;
            session.TranslatorBaseline = RuntimeProcessUtilities.GetProcessIds("JRPG Translator");
            session.OverlayBaseline = RuntimeProcessUtilities.GetProcessIds("overlay");
            bool translatorWasRunning = session.TranslatorBaseline.Count > 0;

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = executable,
                WorkingDirectory = Path.GetDirectoryName(executable) ?? string.Empty,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in RuntimeProcessUtilities.BuildTranslatorArguments(
                translatorWasRunning,
                useBigBoxUi,
                game.TranslatorProfile,
                session.GameContext))
            {
                startInfo.ArgumentList.Add(argument);
            }

            using Process? process = Process.Start(startInfo);
            if (translatorWasRunning)
            {
                RuntimeLog.Write(process == null
                    ? "The running JRPG Translator could not receive the launch settings."
                    : "The launch settings and "
                        + (useBigBoxUi ? "Big Box" : "LaunchBox")
                        + " presentation mode were sent to the running JRPG Translator; it will be left open after the game.");
                return;
            }

            session.TranslatorStartedByPlugin = process != null;
            session.TranslatorProcessId = process?.Id;
            RuntimeLog.Write(process == null
                ? "JRPG Translator could not be started."
                : "JRPG Translator started for " + session.GameTitle
                    + (string.IsNullOrWhiteSpace(game.TranslatorProfile)
                        ? "."
                        : " with Profile '" + game.TranslatorProfile + "'.")
                    + " Presentation mode: " + (useBigBoxUi ? "Big Box." : "desktop."));
        }

        private static void StartJoyToKey(
            PluginConfiguration configuration,
            GameConfiguration game,
            RuntimeSession session)
        {
            string executable = configuration.JoyToKeyExecutable;
            if (!File.Exists(executable))
            {
                RuntimeLog.Write("JoyToKey executable was not found: " + executable);
                return;
            }

            if (string.IsNullOrWhiteSpace(game.JoyToKeyProfile))
            {
                RuntimeLog.Write("No JoyToKey profile was selected for " + session.GameTitle + ".");
                return;
            }

            session.JoyToKeyExecutable = executable;
            session.JoyToKeyProfilesDirectory = configuration.JoyToKeyProfilesDirectory;
            session.JoyToKeyBaseline = RuntimeProcessUtilities.GetProcessIds("JoyToKey");
            session.JoyToKeyWasRunning = session.JoyToKeyBaseline.Count > 0;
            session.PreviousJoyToKeyProfile = RuntimeProcessUtilities.ReadJoyToKeyActiveProfile(
                configuration.JoyToKeyProfilesDirectory);

            using Process? process = RuntimeProcessUtilities.StartWithSingleArgument(
                executable,
                game.JoyToKeyProfile);
            session.JoyToKeyStartProcessId = process?.Id;

            RuntimeLog.Write("JoyToKey profile selected for " + session.GameTitle + ": "
                + game.JoyToKeyProfile + ".");
        }

        private static void StopSession(RuntimeSession session)
        {
            try
            {
                StopJoyToKey(session);
            }
            catch (Exception exception)
            {
                RuntimeLog.Write("JoyToKey cleanup failed: " + exception.Message);
            }

            try
            {
                StopTranslator(session);
            }
            catch (Exception exception)
            {
                RuntimeLog.Write("JRPG Translator cleanup failed: " + exception.Message);
            }
        }

        private static void StopJoyToKey(RuntimeSession session)
        {
            if (string.IsNullOrWhiteSpace(session.JoyToKeyExecutable))
            {
                return;
            }

            if (session.JoyToKeyWasRunning)
            {
                if (!string.IsNullOrWhiteSpace(session.PreviousJoyToKeyProfile))
                {
                    RuntimeProcessUtilities.SwitchJoyToKeyProfile(
                        session.JoyToKeyExecutable,
                        session.PreviousJoyToKeyProfile);
                    RuntimeLog.Write("The previous JoyToKey profile was restored.");
                }

                return;
            }

            HashSet<int> current = RuntimeProcessUtilities.GetProcessIds("JoyToKey");
            current.ExceptWith(session.JoyToKeyBaseline);

            if (current.Count > 0 && !string.IsNullOrWhiteSpace(session.PreviousJoyToKeyProfile))
            {
                RuntimeProcessUtilities.SwitchJoyToKeyProfile(
                    session.JoyToKeyExecutable,
                    session.PreviousJoyToKeyProfile);
                current = RuntimeProcessUtilities.GetProcessIds("JoyToKey");
                current.ExceptWith(session.JoyToKeyBaseline);
            }

            RuntimeProcessUtilities.StopProcesses(current, "JoyToKey");

            if (!string.IsNullOrWhiteSpace(session.PreviousJoyToKeyProfile))
            {
                RuntimeProcessUtilities.WriteJoyToKeyActiveProfile(
                    session.JoyToKeyProfilesDirectory,
                    session.PreviousJoyToKeyProfile);
                RuntimeLog.Write("The previous JoyToKey profile was saved for the next launch.");
            }

            RuntimeLog.Write("The JoyToKey instance started for the game was closed.");
        }

        private static void StopTranslator(RuntimeSession session)
        {
            if (!session.TranslatorStartedByPlugin)
            {
                RuntimeProcessUtilities.SendTranslatorGameContextClear(
                    session.TranslatorExecutable);
                return;
            }

            HashSet<int> translatorProcesses = RuntimeProcessUtilities.GetProcessIds("JRPG Translator");
            translatorProcesses.ExceptWith(session.TranslatorBaseline);
            if (session.TranslatorProcessId.HasValue)
            {
                translatorProcesses.Add(session.TranslatorProcessId.Value);
            }

            RuntimeProcessUtilities.StopProcesses(translatorProcesses, "JRPG Translator");

            HashSet<int> overlayProcesses = RuntimeProcessUtilities.GetProcessIds("overlay");
            overlayProcesses.ExceptWith(session.OverlayBaseline);
            RuntimeProcessUtilities.StopProcesses(overlayProcesses, "overlay");
            RuntimeLog.Write("JRPG Translator was closed after " + session.GameTitle + ".");
        }
    }

    internal sealed class RuntimeSession
    {
        public RuntimeSession(
            string gameId,
            string gameTitle,
            TranslatorGameContext gameContext)
        {
            GameId = gameId;
            GameTitle = gameTitle;
            GameContext = gameContext;
        }

        public string GameId { get; }
        public string GameTitle { get; }
        public TranslatorGameContext GameContext { get; }
        public bool TranslatorStartedByPlugin { get; set; }
        public int? TranslatorProcessId { get; set; }
        public string TranslatorExecutable { get; set; } = string.Empty;
        public HashSet<int> TranslatorBaseline { get; set; } = new HashSet<int>();
        public HashSet<int> OverlayBaseline { get; set; } = new HashSet<int>();
        public string JoyToKeyExecutable { get; set; } = string.Empty;
        public string JoyToKeyProfilesDirectory { get; set; } = string.Empty;
        public bool JoyToKeyWasRunning { get; set; }
        public int? JoyToKeyStartProcessId { get; set; }
        public string PreviousJoyToKeyProfile { get; set; } = string.Empty;
        public HashSet<int> JoyToKeyBaseline { get; set; } = new HashSet<int>();
    }

    public sealed class TranslatorGameContext
    {
        public string Title { get; set; } = string.Empty;
        public string Platform { get; set; } = string.Empty;
        public string BoxArtPath { get; set; } = string.Empty;
        public string ClearLogoPath { get; set; } = string.Empty;
        public string PlatformLogoPath { get; set; } = string.Empty;
        public string PlatformDevicePath { get; set; } = string.Empty;
        public string PlatformDefaultArtPath { get; set; } = string.Empty;
    }

    public static class RuntimeProcessUtilities
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool WritePrivateProfileString(
            string section,
            string key,
            string value,
            string filePath);

        public static IReadOnlyList<string> BuildTranslatorArguments(
            bool translatorWasRunning,
            bool useBigBoxUi,
            string translatorProfile,
            TranslatorGameContext? gameContext = null)
        {
            List<string> arguments = new List<string>
            {
                "--background",
                useBigBoxUi ? "--bigbox-ui" : "--launchbox-ui"
            };
            // Startup overlays belong to the selected Translator Profile (or
            // its current settings when no Profile is selected), not the host.
            if (!string.IsNullOrWhiteSpace(translatorProfile))
            {
                arguments.Add("--profile");
                arguments.Add(translatorProfile);
            }
            if (gameContext != null)
            {
                AddArgumentPair(arguments, "--game-title", gameContext.Title);
                AddArgumentPair(arguments, "--game-platform", gameContext.Platform);
                AddArgumentPair(arguments, "--game-box-art", gameContext.BoxArtPath);
                AddArgumentPair(arguments, "--game-clear-logo", gameContext.ClearLogoPath);
                AddArgumentPair(arguments, "--platform-clear-logo", gameContext.PlatformLogoPath);
                AddArgumentPair(arguments, "--platform-device-image", gameContext.PlatformDevicePath);
                AddArgumentPair(arguments, "--platform-default-art", gameContext.PlatformDefaultArtPath);
            }
            return arguments;
        }

        public static IReadOnlyList<string> BuildTranslatorGameContextClearArguments()
        {
            return new[] { "--background", "--clear-game-context" };
        }

        public static TranslatorGameContext CreateGameContext(IGame game)
        {
            TranslatorGameContext context = new TranslatorGameContext
            {
                Title = game.Title ?? string.Empty,
                Platform = game.Platform ?? string.Empty,
                BoxArtPath = ExistingMediaPath(game.FrontImagePath),
                ClearLogoPath = ExistingMediaPath(game.ClearLogoImagePath),
                PlatformLogoPath = ExistingMediaPath(game.PlatformClearLogoImagePath)
            };

            try
            {
                IPlatform? platform = PluginHelper.DataManager.GetPlatformByName(context.Platform);
                if (platform != null)
                {
                    if (string.IsNullOrWhiteSpace(context.PlatformLogoPath))
                    {
                        context.PlatformLogoPath = ExistingMediaPath(platform.ClearLogoImagePath);
                    }
                    context.PlatformDevicePath = ExistingMediaPath(platform.DeviceImagePath);
                    context.PlatformDefaultArtPath = ExistingMediaPath(platform.DefaultBoxImagePath);
                    if (string.IsNullOrWhiteSpace(context.PlatformDefaultArtPath))
                    {
                        context.PlatformDefaultArtPath = ExistingMediaPath(platform.BackgroundImagePath);
                    }
                }
            }
            catch (Exception exception)
            {
                RuntimeLog.Write("Platform artwork lookup failed: " + exception.Message);
            }

            return context;
        }

        public static void SendTranslatorGameContextClear(string executable)
        {
            if (string.IsNullOrWhiteSpace(executable) || !File.Exists(executable))
            {
                return;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = executable,
                    WorkingDirectory = Path.GetDirectoryName(executable) ?? string.Empty,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                foreach (string argument in BuildTranslatorGameContextClearArguments())
                {
                    startInfo.ArgumentList.Add(argument);
                }
                using Process? process = Process.Start(startInfo);
            }
            catch (Exception exception)
            {
                RuntimeLog.Write("Running-game context could not be cleared: " + exception.Message);
            }
        }

        private static void AddArgumentPair(
            ICollection<string> arguments,
            string name,
            string? value)
        {
            arguments.Add(name);
            arguments.Add(value ?? string.Empty);
        }

        private static string ExistingMediaPath(string? path)
        {
            string candidate = path?.Trim() ?? string.Empty;
            return candidate.Length > 0 && File.Exists(candidate)
                ? candidate
                : string.Empty;
        }

        public static HashSet<int> GetProcessIds(string processName)
        {
            HashSet<int> result = new HashSet<int>();
            try
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    using (process)
                    {
                        result.Add(process.Id);
                    }
                }
            }
            catch
            {
                // Process enumeration can race with process exit; a partial set is safe.
            }

            return result;
        }

        public static string ReadJoyToKeyActiveProfile(string profilesDirectory)
        {
            if (string.IsNullOrWhiteSpace(profilesDirectory))
            {
                return string.Empty;
            }

            string iniFile = Path.Combine(profilesDirectory, "JoyToKey.ini");
            if (!File.Exists(iniFile))
            {
                return string.Empty;
            }

            bool inLastStatus = false;
            try
            {
                foreach (string sourceLine in File.ReadLines(iniFile))
                {
                    string line = sourceLine.Trim();
                    if (line.StartsWith("[", StringComparison.Ordinal)
                        && line.EndsWith("]", StringComparison.Ordinal))
                    {
                        inLastStatus = string.Equals(line, "[LastStatus]", StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    if (!inLastStatus)
                    {
                        continue;
                    }

                    int separator = line.IndexOf('=');
                    if (separator <= 0)
                    {
                        continue;
                    }

                    string key = line.Substring(0, separator).Trim();
                    if (string.Equals(key, "FileName", StringComparison.OrdinalIgnoreCase))
                    {
                        return line.Substring(separator + 1).Trim();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        public static Process? StartWithSingleArgument(string executable, string argument)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = executable,
                WorkingDirectory = Path.GetDirectoryName(executable) ?? string.Empty,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            startInfo.ArgumentList.Add(argument);
            return Process.Start(startInfo);
        }

        public static void SwitchJoyToKeyProfile(string executable, string profile)
        {
            using Process? process = StartWithSingleArgument(executable, profile);
            if (process == null)
            {
                return;
            }

            try
            {
                process.WaitForExit(2000);
            }
            catch (InvalidOperationException)
            {
            }

            // Give the running instance a brief moment to process the switch
            // message before an owned instance is stopped.
            Thread.Sleep(250);
        }

        public static bool WriteJoyToKeyActiveProfile(string profilesDirectory, string profile)
        {
            if (string.IsNullOrWhiteSpace(profilesDirectory)
                || string.IsNullOrWhiteSpace(profile))
            {
                return false;
            }

            string iniFile = Path.Combine(profilesDirectory, "JoyToKey.ini");
            if (!File.Exists(iniFile))
            {
                return false;
            }

            return WritePrivateProfileString("LastStatus", "FileName", profile, iniFile);
        }

        public static void StopProcesses(IEnumerable<int> processIds, params string[] expectedProcessNames)
        {
            foreach (int processId in processIds.Distinct())
            {
                try
                {
                    using (Process process = Process.GetProcessById(processId))
                    {
                        if (!MatchesExpectedProcessName(process, expectedProcessNames))
                        {
                            // PIDs can be reused after a plugin-started process exits.
                            // Never terminate a different process that inherited the ID.
                            RuntimeLog.Write("Skipped cleanup for reused process ID " + processId + ".");
                            continue;
                        }

                        process.Kill(true);
                        process.WaitForExit(1500);
                    }
                }
                catch (ArgumentException)
                {
                    // The process already exited.
                }
                catch (InvalidOperationException)
                {
                    // The process already exited.
                }
                catch (System.ComponentModel.Win32Exception)
                {
                    // A process we do not own could not be opened; leave it alone.
                }
            }
        }

        private static bool MatchesExpectedProcessName(Process process, IReadOnlyCollection<string> expectedNames)
        {
            if (expectedNames.Count == 0)
            {
                return false;
            }

            string processName = process.ProcessName;
            return expectedNames.Any(expected =>
                !string.IsNullOrWhiteSpace(expected)
                && string.Equals(processName, expected, StringComparison.OrdinalIgnoreCase));
        }
    }

    internal static class RuntimeLog
    {
        private static readonly object Sync = new object();

        public static void Write(string message)
        {
            try
            {
                lock (Sync)
                {
                    Directory.CreateDirectory(PluginPaths.DataDirectory);
                    File.AppendAllText(
                        Path.Combine(PluginPaths.DataDirectory, "runtime.log"),
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message
                            + Environment.NewLine);
                }
            }
            catch
            {
                // Runtime logging must never interfere with game launch.
            }
        }
    }
}
