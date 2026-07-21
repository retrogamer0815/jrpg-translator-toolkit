using System;
using System.Drawing;
using System.Windows;
using Unbroken.LaunchBox.Plugins;
using Unbroken.LaunchBox.Plugins.Data;

namespace JrpgTranslator.LaunchBox
{
    public sealed class GameSetupMenuItem : IGameMenuItemPlugin
    {
        public bool SupportsMultipleGames { get { return false; } }
        public string Caption { get { return "JRPG Translator Setup..."; } }
        public Image IconImage { get { return null!; } }
        public bool ShowInLaunchBox { get { return true; } }
        public bool ShowInBigBox { get { return true; } }

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
