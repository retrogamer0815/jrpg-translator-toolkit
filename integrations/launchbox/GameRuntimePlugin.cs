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

            RuntimeCoordinator.Start(game.Id, game.Title);
        }

        public void OnAfterGameLaunched(IGame? game, IAdditionalApplication? app, IEmulator? emulator)
        {
        }

        public void OnGameExited()
        {
            RuntimeCoordinator.StopActiveSession();
        }
    }

    public sealed class RuntimeShutdownPlugin : ISystemEventsPlugin
    {
        public void OnEventRaised(string eventType)
        {
            if (string.Equals(eventType, SystemEventTypes.LaunchBoxShutdownBeginning, StringComparison.Ordinal)
                || string.Equals(eventType, SystemEventTypes.BigBoxShutdownBeginning, StringComparison.Ordinal))
            {
                RuntimeCoordinator.StopActiveSession();
            }
        }
    }

    internal static class RuntimeCoordinator
    {
        private static readonly object Sync = new object();
        private static RuntimeSession? _activeSession;

        public static void Start(string gameId, string gameTitle)
        {
            lock (Sync)
            {
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
                    if (!game.TranslatorEnabled && !game.JoyToKeyEnabled)
                    {
                        return;
                    }

                    RuntimeSession session = new RuntimeSession(gameId, gameTitle);
                    _activeSession = session;

                    if (game.TranslatorEnabled)
                    {
                        StartTranslator(configuration, game, session);
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
            lock (Sync)
            {
                if (_activeSession == null)
                {
                    return;
                }

                RuntimeSession session = _activeSession;
                _activeSession = null;
                StopSession(session);
            }
        }

        private static void StartTranslator(
            PluginConfiguration configuration,
            GameConfiguration game,
            RuntimeSession session)
        {
            string executable = PluginPaths.ResolveTranslatorExecutable(configuration);
            if (!File.Exists(executable))
            {
                RuntimeLog.Write("JRPG Translator executable was not found: " + executable);
                return;
            }

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
            startInfo.ArgumentList.Add("--background");
            if (!translatorWasRunning)
            {
                startInfo.ArgumentList.Add("--open-translator");
            }
            if (!string.IsNullOrWhiteSpace(game.TranslatorProfile))
            {
                startInfo.ArgumentList.Add("--profile");
                startInfo.ArgumentList.Add(game.TranslatorProfile);
            }

            using Process? process = Process.Start(startInfo);
            if (translatorWasRunning)
            {
                RuntimeLog.Write(process == null
                    ? "The running JRPG Translator could not receive the selected profile."
                    : "The selected Profile was sent to the running JRPG Translator; it will be left open after the game.");
                return;
            }

            session.TranslatorStartedByPlugin = process != null;
            session.TranslatorProcessId = process?.Id;
            RuntimeLog.Write(process == null
                ? "JRPG Translator could not be started."
                : "JRPG Translator started for " + session.GameTitle
                    + (string.IsNullOrWhiteSpace(game.TranslatorProfile)
                        ? "."
                        : " with Profile '" + game.TranslatorProfile + "'."));
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

            RuntimeProcessUtilities.StopProcesses(current);

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
                return;
            }

            HashSet<int> translatorProcesses = RuntimeProcessUtilities.GetProcessIds("JRPG Translator");
            translatorProcesses.ExceptWith(session.TranslatorBaseline);
            if (session.TranslatorProcessId.HasValue)
            {
                translatorProcesses.Add(session.TranslatorProcessId.Value);
            }

            RuntimeProcessUtilities.StopProcesses(translatorProcesses);

            HashSet<int> overlayProcesses = RuntimeProcessUtilities.GetProcessIds("overlay");
            overlayProcesses.ExceptWith(session.OverlayBaseline);
            RuntimeProcessUtilities.StopProcesses(overlayProcesses);
            RuntimeLog.Write("JRPG Translator was closed after " + session.GameTitle + ".");
        }
    }

    internal sealed class RuntimeSession
    {
        public RuntimeSession(string gameId, string gameTitle)
        {
            GameId = gameId;
            GameTitle = gameTitle;
        }

        public string GameId { get; }
        public string GameTitle { get; }
        public bool TranslatorStartedByPlugin { get; set; }
        public int? TranslatorProcessId { get; set; }
        public HashSet<int> TranslatorBaseline { get; set; } = new HashSet<int>();
        public HashSet<int> OverlayBaseline { get; set; } = new HashSet<int>();
        public string JoyToKeyExecutable { get; set; } = string.Empty;
        public string JoyToKeyProfilesDirectory { get; set; } = string.Empty;
        public bool JoyToKeyWasRunning { get; set; }
        public int? JoyToKeyStartProcessId { get; set; }
        public string PreviousJoyToKeyProfile { get; set; } = string.Empty;
        public HashSet<int> JoyToKeyBaseline { get; set; } = new HashSet<int>();
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

        public static void StopProcesses(IEnumerable<int> processIds)
        {
            foreach (int processId in processIds.Distinct())
            {
                try
                {
                    using (Process process = Process.GetProcessById(processId))
                    {
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
