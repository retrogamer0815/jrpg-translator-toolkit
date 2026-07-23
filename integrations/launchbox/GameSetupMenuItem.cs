using System;
using System.Drawing;
using System.IO;
using System.Windows;
using Unbroken.LaunchBox.Plugins;
using Unbroken.LaunchBox.Plugins.Data;

namespace JrpgTranslator.LaunchBox
{
    public sealed class GameSetupMenuItem : IGameMenuItemPlugin
    {
        private static readonly object MenuIconLock = new object();
        private static Stream? _menuIconStream;
        private static Image? _menuIcon;

        public bool SupportsMultipleGames { get { return false; } }
        public string Caption { get { return "JRPG Translator Setup..."; } }
        public Image IconImage { get { return LoadMenuIcon(); } }
        public bool ShowInLaunchBox { get { return true; } }
        public bool ShowInBigBox { get { return true; } }

        private static Image LoadMenuIcon()
        {
            lock (MenuIconLock)
            {
                if (_menuIcon != null)
                {
                    return _menuIcon;
                }

                try
                {
                    _menuIconStream = typeof(GameSetupMenuItem).Assembly.GetManifestResourceStream(
                        "JrpgTranslator.LaunchBox.Assets.menu-icon.png");
                    if (_menuIconStream == null)
                    {
                        return null!;
                    }

                    _menuIcon = Image.FromStream(_menuIconStream);
                    return _menuIcon;
                }
                catch
                {
                    _menuIconStream?.Dispose();
                    _menuIconStream = null;
                    return null!;
                }
            }
        }

        public bool GetIsValidForGame(IGame selectedGame)
        {
            return selectedGame != null;
        }

        public bool GetIsValidForGames(IGame[] selectedGames)
        {
            return false;
        }

        public void OnSelected(IGame selectedGame)
        {
            if (selectedGame == null)
            {
                return;
            }

            try
            {
                PluginConfiguration configuration = ConfigurationStore.Load();
                GameConfiguration game = configuration.GetGame(selectedGame.Id, selectedGame.Title).Clone();
                GameSetupWindow dialog = new GameSetupWindow(configuration, game);

                if (Application.Current?.MainWindow != null
                    && Application.Current.MainWindow != dialog
                    && Application.Current.MainWindow.IsVisible)
                {
                    dialog.Owner = Application.Current.MainWindow;
                }

                if (dialog.ShowDialog() == true)
                {
                    configuration.UpsertGame(dialog.Result);
                    ConfigurationStore.Save(configuration);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "JRPG Translator setup could not be opened.\n\n" + exception.Message,
                    "JRPG Translator Integration",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        public void OnSelected(IGame[] selectedGames)
        {
        }
    }
}
