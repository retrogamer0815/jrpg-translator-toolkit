using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace JrpgTranslator.LaunchBox
{
    [Serializable]
    public sealed class PluginConfiguration
    {
        public string TranslatorExecutable { get; set; } = @"Apps\JRPG Translator\JRPG Translator.exe";
        public string JoyToKeyExecutable { get; set; } = string.Empty;
        public string JoyToKeyProfilesDirectory { get; set; } = string.Empty;
        public List<GameConfiguration> Games { get; set; } = new List<GameConfiguration>();

        public GameConfiguration GetGame(string gameId, string title)
        {
            GameConfiguration? game = Games.FirstOrDefault(
                item => string.Equals(item.GameId, gameId, StringComparison.OrdinalIgnoreCase));

            if (game == null)
            {
                game = new GameConfiguration { GameId = gameId, GameTitle = title };
            }
            else
            {
                game.GameTitle = title;
            }

            return game;
        }

        public void UpsertGame(GameConfiguration game)
        {
            Games.RemoveAll(item => string.Equals(item.GameId, game.GameId, StringComparison.OrdinalIgnoreCase));
            Games.Add(game);
            Games.Sort((left, right) => StringComparer.CurrentCultureIgnoreCase.Compare(left.GameTitle, right.GameTitle));
        }
    }

    [Serializable]
    public sealed class GameConfiguration
    {
        public string GameId { get; set; } = string.Empty;
        public string GameTitle { get; set; } = string.Empty;
        public bool TranslatorEnabled { get; set; }
        public bool JoyToKeyEnabled { get; set; }
        public string JoyToKeyProfile { get; set; } = string.Empty;

        public GameConfiguration Clone()
        {
            return new GameConfiguration
            {
                GameId = GameId,
                GameTitle = GameTitle,
                TranslatorEnabled = TranslatorEnabled,
                JoyToKeyEnabled = JoyToKeyEnabled,
                JoyToKeyProfile = JoyToKeyProfile
            };
        }
    }

    public static class ConfigurationStore
    {
        private static readonly object Sync = new object();

        public static PluginConfiguration Load()
        {
            lock (Sync)
            {
                PluginConfiguration configuration;
                if (!File.Exists(PluginPaths.ConfigurationFile))
                {
                    configuration = new PluginConfiguration();
                }
                else
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(PluginConfiguration));
                    using (FileStream stream = File.OpenRead(PluginPaths.ConfigurationFile))
                    {
                        configuration = (PluginConfiguration?)serializer.Deserialize(stream)
                            ?? new PluginConfiguration();
                    }
                }

                PluginPaths.PopulateDetectedPaths(configuration);
                return configuration;
            }
        }

        public static void Save(PluginConfiguration configuration)
        {
            lock (Sync)
            {
                Directory.CreateDirectory(PluginPaths.DataDirectory);
                string temporaryFile = PluginPaths.ConfigurationFile + ".tmp";
                XmlSerializer serializer = new XmlSerializer(typeof(PluginConfiguration));

                using (FileStream stream = File.Create(temporaryFile))
                {
                    serializer.Serialize(stream, configuration);
                }

                if (File.Exists(PluginPaths.ConfigurationFile))
                {
                    File.Replace(temporaryFile, PluginPaths.ConfigurationFile, null);
                }
                else
                {
                    File.Move(temporaryFile, PluginPaths.ConfigurationFile);
                }
            }
        }
    }

    public static class PluginPaths
    {
        public static string PluginDirectory
        {
            get
            {
                string? location = Assembly.GetExecutingAssembly().Location;
                return Path.GetDirectoryName(location) ?? AppContext.BaseDirectory;
            }
        }

        public static string DataDirectory { get { return Path.Combine(PluginDirectory, "PluginData"); } }
        public static string ConfigurationFile { get { return Path.Combine(DataDirectory, "config.xml"); } }

        public static string LaunchBoxRoot
        {
            get
            {
                DirectoryInfo? directory = new DirectoryInfo(PluginDirectory);
                for (int depth = 0; directory != null && depth < 6; depth++, directory = directory.Parent)
                {
                    if (File.Exists(Path.Combine(directory.FullName, "LaunchBox.exe")))
                    {
                        return directory.FullName;
                    }
                }

                return string.Empty;
            }
        }

        public static string ResolveTranslatorExecutable(PluginConfiguration configuration)
        {
            if (Path.IsPathRooted(configuration.TranslatorExecutable))
            {
                return configuration.TranslatorExecutable;
            }

            return string.IsNullOrWhiteSpace(LaunchBoxRoot)
                ? configuration.TranslatorExecutable
                : Path.Combine(LaunchBoxRoot, configuration.TranslatorExecutable);
        }

        public static void PopulateDetectedPaths(PluginConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(configuration.TranslatorExecutable))
            {
                configuration.TranslatorExecutable = @"Apps\JRPG Translator\JRPG Translator.exe";
            }
            else if (string.Equals(
                configuration.TranslatorExecutable,
                @"Apps\Apps\JRPG Translator\JRPG Translator.exe",
                StringComparison.OrdinalIgnoreCase))
            {
                string correctedRelativePath = @"Apps\JRPG Translator\JRPG Translator.exe";
                string correctedFullPath = string.IsNullOrWhiteSpace(LaunchBoxRoot)
                    ? correctedRelativePath
                    : Path.Combine(LaunchBoxRoot, correctedRelativePath);
                if (File.Exists(correctedFullPath))
                {
                    configuration.TranslatorExecutable = correctedRelativePath;
                }
            }

            if (string.IsNullOrWhiteSpace(configuration.JoyToKeyExecutable)
                || !File.Exists(configuration.JoyToKeyExecutable))
            {
                configuration.JoyToKeyExecutable = DetectJoyToKeyExecutable();
            }

            if (string.IsNullOrWhiteSpace(configuration.JoyToKeyProfilesDirectory)
                || !Directory.Exists(configuration.JoyToKeyProfilesDirectory))
            {
                configuration.JoyToKeyProfilesDirectory = DetectJoyToKeyProfilesDirectory(
                    configuration.JoyToKeyExecutable);
            }
        }

        private static string DetectJoyToKeyExecutable()
        {
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            string installed = Path.Combine(programFilesX86, "JoyToKey", "JoyToKey.exe");
            if (File.Exists(installed))
            {
                return installed;
            }

            string portable = string.IsNullOrWhiteSpace(LaunchBoxRoot)
                ? string.Empty
                : Path.Combine(LaunchBoxRoot, "Apps", "JoyToKey", "JoyToKey.exe");
            return File.Exists(portable) ? portable : string.Empty;
        }

        private static string DetectJoyToKeyProfilesDirectory(string joyToKeyExecutable)
        {
            List<string> candidates = new List<string>();

            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Software\JoyToKey"))
                {
                    string? configured = key?.GetValue("DataDir") as string;
                    if (!string.IsNullOrWhiteSpace(configured))
                    {
                        candidates.Add(Environment.ExpandEnvironmentVariables(configured));
                    }
                }
            }
            catch
            {
                // Registry detection is optional; common folders are checked below.
            }

            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            candidates.Add(Path.Combine(documents, "JoyToKey"));
            candidates.Add(Path.Combine(userProfile, "My Drive", "Documents", "JoyToKey"));

            if (!string.IsNullOrWhiteSpace(joyToKeyExecutable))
            {
                string? executableDirectory = Path.GetDirectoryName(joyToKeyExecutable);
                if (!string.IsNullOrWhiteSpace(executableDirectory))
                {
                    candidates.Add(executableDirectory);
                }
            }

            return candidates
                .Where(Directory.Exists)
                .OrderByDescending(path => Directory.EnumerateFiles(path, "*.cfg").Any())
                .FirstOrDefault() ?? string.Empty;
        }
    }

    public static class JoyToKeyProfileDiscovery
    {
        public static IReadOnlyList<string> GetProfiles(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                return Array.Empty<string>();
            }

            return Directory.EnumerateFiles(directory, "*.cfg", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileNameWithoutExtension)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToArray()!;
        }
    }
}
