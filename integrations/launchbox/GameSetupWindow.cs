using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;

namespace JrpgTranslator.LaunchBox
{
    public sealed class GameSetupWindow : Window
    {
        private static readonly Brush WindowBackground = BrushFrom("#202124");
        private static readonly Brush PanelBackground = BrushFrom("#292A2D");
        private static readonly Brush ControlBackground = BrushFrom("#333438");
        private static readonly Brush PrimaryForeground = BrushFrom("#F3F4F6");
        private static readonly Brush MutedForeground = BrushFrom("#AEB3BC");
        private static readonly Brush ControlBorderBrush = BrushFrom("#686B72");

        private readonly PluginConfiguration _configuration;
        private readonly CheckBox _translatorEnabled;
        private readonly CheckBox _joyToKeyEnabled;
        private readonly ComboBox _translatorProfile;
        private readonly Button _refreshTranslatorProfiles;
        private readonly ComboBox _joyToKeyProfile;
        private readonly Button _refreshJoyToKeyProfiles;
        private readonly Button _browseTranslator;
        private readonly Button _browseJoyToKey;
        private readonly Button _browseProfiles;
        private readonly Button _save;
        private readonly Button _cancel;
        private readonly TextBlock _translatorProfileStatus;
        private readonly TextBlock _joyToKeyProfileStatus;
        private readonly TextBlock _readiness;
        private readonly List<Control[]> _focusRows;
        private readonly DispatcherTimer _controllerTimer;
        private readonly bool _nativeControllerNavigationEnabled;
        private readonly Dictionary<ControllerNavigationCommand, long> _keyboardMirrorGraceUntil = new();
        private readonly Dictionary<ControllerNavigationCommand, long> _lastNativeNavigationAt = new();
        private ControllerNavigationState _controllerPreviousState;
        private ControllerNavigationCommand? _heldControllerDirection;
        private long _nextControllerRepeatAt;
        private bool _controllerBaselineReady;
        private ComboBox? _guardedControllerProfile;
        private long _guardedControllerProfileUntil;

        public GameConfiguration Result { get; private set; }

        public GameSetupWindow(PluginConfiguration configuration, GameConfiguration game)
        {
            _configuration = configuration;
            Result = game;
            _nativeControllerNavigationEnabled = IsBigBoxHost();

            Title = "JRPG Translator Setup";
            Width = 760;
            Height = 780;
            MinWidth = 680;
            MinHeight = 690;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            ShowInTaskbar = false;
            Background = WindowBackground;
            Foreground = PrimaryForeground;
            Cursor = Cursors.Arrow;
            ForceCursor = true;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 18;

            Grid root = new Grid { Margin = new Thickness(28) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            TextBlock heading = new TextBlock
            {
                Text = game.GameTitle,
                FontSize = 25,
                FontWeight = FontWeights.SemiBold,
                Foreground = PrimaryForeground,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(0, 0, 0, 8)
            };
            root.Children.Add(heading);

            TextBlock introduction = new TextBlock
            {
                Text = "Choose what should be prepared automatically whenever this game is launched.",
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 22)
            };
            Grid.SetRow(introduction, 1);
            root.Children.Add(introduction);

            StackPanel options = new StackPanel
            {
                Background = PanelBackground,
                Margin = new Thickness(0, 0, 0, 18)
            };
            Grid.SetRow(options, 2);
            root.Children.Add(options);

            _translatorEnabled = MakeCheckBox("Use JRPG Translator with this game", game.TranslatorEnabled);
            _translatorEnabled.Margin = new Thickness(18, 17, 18, 10);
            _translatorEnabled.Checked += (_, _) => UpdateTranslatorControls();
            _translatorEnabled.Unchecked += (_, _) => UpdateTranslatorControls();
            options.Children.Add(_translatorEnabled);

            Grid translatorProfileRow = MakeProfileRow(out _translatorProfile, out _refreshTranslatorProfiles);
            _translatorProfile.DropDownClosed += HandleProfileDropDownClosed;
            _refreshTranslatorProfiles.Click += (_, _) => RefreshTranslatorProfiles(_translatorProfile.Text);
            options.Children.Add(translatorProfileRow);

            _translatorProfileStatus = new TextBlock
            {
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(54, 0, 18, 12)
            };
            options.Children.Add(_translatorProfileStatus);

            _joyToKeyEnabled = MakeCheckBox("Use JoyToKey with this game", game.JoyToKeyEnabled);
            _joyToKeyEnabled.Margin = new Thickness(18, 10, 18, 14);
            _joyToKeyEnabled.Checked += (_, _) => UpdateJoyToKeyControls();
            _joyToKeyEnabled.Unchecked += (_, _) => UpdateJoyToKeyControls();
            options.Children.Add(_joyToKeyEnabled);

            Grid joyToKeyProfileRow = MakeProfileRow(out _joyToKeyProfile, out _refreshJoyToKeyProfiles);
            _joyToKeyProfile.DropDownClosed += HandleProfileDropDownClosed;
            _refreshJoyToKeyProfiles.Click += (_, _) => RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);
            options.Children.Add(joyToKeyProfileRow);

            _joyToKeyProfileStatus = new TextBlock
            {
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(54, 0, 18, 12)
            };
            options.Children.Add(_joyToKeyProfileStatus);

            TextBlock locationHeading = new TextBlock
            {
                Text = "Application locations",
                Foreground = MutedForeground,
                Margin = new Thickness(18, 4, 18, 8)
            };
            options.Children.Add(locationHeading);

            WrapPanel locationButtons = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(18, 0, 18, 16)
            };
            options.Children.Add(locationButtons);

            _browseTranslator = MakeButton("Translator EXE...", 150);
            _browseTranslator.Click += BrowseTranslatorClicked;
            locationButtons.Children.Add(_browseTranslator);

            _browseJoyToKey = MakeButton("JoyToKey EXE...", 150);
            _browseJoyToKey.Margin = new Thickness(10, 0, 0, 0);
            _browseJoyToKey.Click += BrowseJoyToKeyClicked;
            locationButtons.Children.Add(_browseJoyToKey);

            _browseProfiles = MakeButton("JoyToKey Profiles...", 178);
            _browseProfiles.Margin = new Thickness(10, 0, 0, 0);
            _browseProfiles.Click += BrowseProfilesClicked;
            locationButtons.Children.Add(_browseProfiles);

            _readiness = new TextBlock
            {
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12),
                Text = BuildReadinessText()
            };
            Grid.SetRow(_readiness, 3);
            root.Children.Add(_readiness);

            TextBlock stageNotice = new TextBlock
            {
                Text = "JRPG Translator and the selected JoyToKey profile are prepared automatically for this game, then closed or restored after the game exits.",
                Foreground = MutedForeground,
                FontStyle = FontStyles.Italic,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Top
            };
            Grid.SetRow(stageNotice, 4);
            root.Children.Add(stageNotice);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 22, 0, 0)
            };
            Grid.SetRow(buttons, 5);
            root.Children.Add(buttons);

            _save = MakeButton("Save", 150);
            _save.IsDefault = true;
            _save.Click += SaveClicked;
            buttons.Children.Add(_save);

            _cancel = MakeButton("Cancel", 150);
            _cancel.IsCancel = true;
            _cancel.Margin = new Thickness(12, 0, 0, 0);
            _cancel.Click += (_, _) => { DialogResult = false; };
            buttons.Children.Add(_cancel);

            _focusRows = new List<Control[]>
            {
                new Control[] { _translatorEnabled },
                new Control[] { _translatorProfile, _refreshTranslatorProfiles },
                new Control[] { _joyToKeyEnabled },
                new Control[] { _joyToKeyProfile, _refreshJoyToKeyProfiles },
                new Control[] { _browseTranslator, _browseJoyToKey, _browseProfiles },
                new Control[] { _save, _cancel }
            };

            _controllerTimer = new DispatcherTimer(DispatcherPriority.Input)
            {
                Interval = TimeSpan.FromMilliseconds(20)
            };
            _controllerTimer.Tick += HandleControllerTick;

            PreviewKeyDown += HandlePreviewKeyDown;
            PreviewKeyUp += HandlePreviewKeyUp;
            Loaded += (_, _) =>
            {
                // LaunchBox can keep its busy cursor active while a synchronous
                // plugin command is open. This dialog has no background work.
                Mouse.OverrideCursor = null;
                Cursor = Cursors.Arrow;
                ForceCursor = true;
                _translatorEnabled.Focus();
                _controllerTimer.Start();
            };
            Activated += (_, _) => ResetControllerNavigation();
            Deactivated += (_, _) => ResetControllerNavigation();
            Closed += (_, _) => _controllerTimer.Stop();

            RefreshTranslatorProfiles(game.TranslatorProfile);
            RefreshJoyToKeyProfiles(game.JoyToKeyProfile);
            UpdateTranslatorControls();
            UpdateJoyToKeyControls();
        }

        private void RefreshTranslatorProfiles(string preferredProfile)
        {
            IReadOnlyList<string> profiles = TranslatorProfileDiscovery.GetProfiles(_configuration);
            _translatorProfile.ItemsSource = profiles;

            string? selected = profiles.FirstOrDefault(
                value => string.Equals(value, preferredProfile, StringComparison.OrdinalIgnoreCase));
            _translatorProfile.SelectedItem = selected;

            string directory = TranslatorProfileDiscovery.GetProfilesDirectory(_configuration);
            _translatorProfileStatus.Text = profiles.Count == 0
                ? "No Profiles were found in " + DisplayPath(directory)
                : profiles.Count + (profiles.Count == 1 ? " Profile" : " Profiles")
                    + " found in " + DisplayPath(directory);
        }

        private void RefreshJoyToKeyProfiles(string preferredProfile)
        {
            IReadOnlyList<string> profiles = JoyToKeyProfileDiscovery.GetProfiles(
                _configuration.JoyToKeyProfilesDirectory);
            _joyToKeyProfile.ItemsSource = profiles;

            string? selected = profiles.FirstOrDefault(
                value => string.Equals(value, preferredProfile, StringComparison.OrdinalIgnoreCase));
            _joyToKeyProfile.SelectedItem = selected ?? profiles.FirstOrDefault();

            _joyToKeyProfileStatus.Text = profiles.Count == 0
                ? "No .cfg profiles were found in " + DisplayPath(_configuration.JoyToKeyProfilesDirectory)
                : profiles.Count + (profiles.Count == 1 ? " profile" : " profiles")
                    + " found in " + DisplayPath(_configuration.JoyToKeyProfilesDirectory);
        }

        private string BuildReadinessText()
        {
            string translator = PluginPaths.ResolveTranslatorExecutable(_configuration);
            string translatorState = File.Exists(translator) ? "ready" : "not found";
            string joyToKeyState = File.Exists(_configuration.JoyToKeyExecutable) ? "ready" : "not found";
            return "JRPG Translator: " + translatorState + "    |    JoyToKey: " + joyToKeyState;
        }

        private static string DisplayPath(string path)
        {
            return string.IsNullOrWhiteSpace(path) ? "an undetected profile folder." : path;
        }

        private void UpdateJoyToKeyControls()
        {
            bool enabled = _joyToKeyEnabled.IsChecked == true;
            _joyToKeyProfile.IsEnabled = enabled;
            _refreshJoyToKeyProfiles.IsEnabled = enabled;
        }

        private void UpdateTranslatorControls()
        {
            bool enabled = _translatorEnabled.IsChecked == true;
            _translatorProfile.IsEnabled = enabled;
            _refreshTranslatorProfiles.IsEnabled = enabled;

            if (enabled && _translatorProfile.SelectedItem == null && _translatorProfile.Items.Count > 0)
            {
                _translatorProfile.SelectedIndex = 0;
            }
        }

        private void BrowseTranslatorClicked(object sender, RoutedEventArgs e)
        {
            string current = PluginPaths.ResolveTranslatorExecutable(_configuration);
            string? selected = _nativeControllerNavigationEnabled
                ? BigBoxPathBrowserWindow.BrowseExecutable(
                    this,
                    "Select JRPG Translator.exe",
                    current)
                : BrowseForExecutable(
                    "Select JRPG Translator.exe",
                    current);
            if (string.IsNullOrWhiteSpace(selected))
            {
                return;
            }

            _configuration.TranslatorExecutable = MakeLaunchBoxRelative(selected);
            RefreshTranslatorProfiles(_translatorProfile.Text);
            UpdatePathStatus();
        }

        private void BrowseJoyToKeyClicked(object sender, RoutedEventArgs e)
        {
            string? selected = _nativeControllerNavigationEnabled
                ? BigBoxPathBrowserWindow.BrowseExecutable(
                    this,
                    "Select JoyToKey.exe",
                    _configuration.JoyToKeyExecutable)
                : BrowseForExecutable(
                    "Select JoyToKey.exe",
                    _configuration.JoyToKeyExecutable);
            if (string.IsNullOrWhiteSpace(selected))
            {
                return;
            }

            _configuration.JoyToKeyExecutable = selected;
            UpdatePathStatus();
        }

        private void BrowseProfilesClicked(object sender, RoutedEventArgs e)
        {
            if (_nativeControllerNavigationEnabled)
            {
                string? selected = BigBoxPathBrowserWindow.BrowseFolder(
                    this,
                    "Select the JoyToKey profiles folder",
                    _configuration.JoyToKeyProfilesDirectory);
                if (string.IsNullOrWhiteSpace(selected))
                {
                    return;
                }

                _configuration.JoyToKeyProfilesDirectory = selected;
                RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);
                UpdatePathStatus();
                return;
            }

            OpenFolderDialog dialog = new OpenFolderDialog
            {
                Title = "Select the JoyToKey profiles folder",
                Multiselect = false
            };

            if (Directory.Exists(_configuration.JoyToKeyProfilesDirectory))
            {
                dialog.InitialDirectory = _configuration.JoyToKeyProfilesDirectory;
            }

            Mouse.OverrideCursor = null;
            bool? accepted;
            using (new XInputDialogCancelScope(this))
            {
                accepted = dialog.ShowDialog(this);
            }
            if (accepted != true || string.IsNullOrWhiteSpace(dialog.FolderName))
            {
                return;
            }

            _configuration.JoyToKeyProfilesDirectory = dialog.FolderName;
            RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);
            UpdatePathStatus();
        }

        private string? BrowseForExecutable(string title, string currentPath)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = title,
                Filter = "Applications (*.exe)|*.exe|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (File.Exists(currentPath))
            {
                dialog.InitialDirectory = Path.GetDirectoryName(currentPath);
                dialog.FileName = Path.GetFileName(currentPath);
            }
            else
            {
                string? directory = Path.GetDirectoryName(currentPath);
                if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                {
                    dialog.InitialDirectory = directory;
                }
            }

            Mouse.OverrideCursor = null;
            bool? accepted;
            using (new XInputDialogCancelScope(this))
            {
                accepted = dialog.ShowDialog(this);
            }
            return accepted == true ? dialog.FileName : null;
        }

        private static string MakeLaunchBoxRelative(string selectedPath)
        {
            if (string.IsNullOrWhiteSpace(PluginPaths.LaunchBoxRoot))
            {
                return selectedPath;
            }

            string relative = Path.GetRelativePath(PluginPaths.LaunchBoxRoot, selectedPath);
            return relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                || string.Equals(relative, "..", StringComparison.Ordinal)
                ? selectedPath
                : relative;
        }

        private void UpdatePathStatus()
        {
            _readiness.Text = BuildReadinessText();
        }

        private void SaveClicked(object sender, RoutedEventArgs e)
        {
            if (_translatorEnabled.IsChecked == true
                && _translatorProfile.Items.Count > 0
                && _translatorProfile.SelectedItem == null)
            {
                MessageBox.Show(
                    this,
                    "Select a Profile, or turn off JRPG Translator for this game.",
                    "JRPG Translator Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                _translatorProfile.Focus();
                return;
            }

            if (_joyToKeyEnabled.IsChecked == true && _joyToKeyProfile.SelectedItem == null)
            {
                MessageBox.Show(
                    this,
                    "Select a JoyToKey profile, or turn off JoyToKey profile switching for this game.",
                    "JRPG Translator Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                _joyToKeyProfile.Focus();
                return;
            }

            Result.TranslatorEnabled = _translatorEnabled.IsChecked == true;
            Result.TranslatorProfile = _translatorProfile.SelectedItem as string ?? string.Empty;
            Result.JoyToKeyEnabled = _joyToKeyEnabled.IsChecked == true;
            Result.JoyToKeyProfile = _joyToKeyProfile.SelectedItem as string ?? string.Empty;
            DialogResult = true;
        }

        private void HandlePreviewKeyDown(object sender, KeyEventArgs e)
        {
            ControllerNavigationCommand? mirroredCommand = ControllerCommandForKey(e.Key);
            if (mirroredCommand.HasValue
                && IsMirroredControllerKey(mirroredCommand.Value))
            {
                e.Handled = true;
                return;
            }

            ComboBox? openProfile = GetOpenProfile();
            if (openProfile != null)
            {
                if (e.Key == Key.Enter || e.Key == Key.Escape)
                {
                    CloseProfileDropDown(openProfile);
                    e.Handled = true;
                }
                return;
            }

            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                MoveFocusVertical(e.Key == Key.Down ? 1 : -1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Left || e.Key == Key.Right)
            {
                MoveFocusHorizontal(e.Key == Key.Right ? 1 : -1);
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Enter && e.Key != Key.Space)
            {
                return;
            }

            e.Handled = ActivateFocusedControl();
        }

        private void HandlePreviewKeyUp(object sender, KeyEventArgs e)
        {
            ControllerNavigationCommand? mirroredCommand = ControllerCommandForKey(e.Key);
            if (mirroredCommand.HasValue
                && IsMirroredControllerKey(mirroredCommand.Value))
            {
                // LaunchBox/Big Box or a controller mapper can emit a matching
                // key-up after the native XInput action. Some WPF controls act
                // on key release, so consume both halves of the mirrored press.
                e.Handled = true;
            }
        }

        private static ControllerNavigationCommand? ControllerCommandForKey(Key key)
        {
            return key switch
            {
                Key.Up => ControllerNavigationCommand.Up,
                Key.Down => ControllerNavigationCommand.Down,
                Key.Left => ControllerNavigationCommand.Left,
                Key.Right => ControllerNavigationCommand.Right,
                Key.Enter => ControllerNavigationCommand.Activate,
                Key.Space => ControllerNavigationCommand.Activate,
                Key.Escape => ControllerNavigationCommand.Cancel,
                _ => null
            };
        }

        private bool IsMirroredControllerKey(ControllerNavigationCommand command)
        {
            const long releaseGraceMilliseconds = 180;
            const long nativeActionGraceMilliseconds = 650;
            long now = Environment.TickCount64;
            if (_lastNativeNavigationAt.TryGetValue(command, out long nativeActionAt)
                && now - nativeActionAt <= nativeActionGraceMilliseconds)
            {
                return true;
            }

            if (XInputController.TryReadNavigationState(out ControllerNavigationState state)
                && state.IsPressed(command))
            {
                _keyboardMirrorGraceUntil[command] = now + releaseGraceMilliseconds;
                return true;
            }

            return _keyboardMirrorGraceUntil.TryGetValue(command, out long graceUntil)
                && now <= graceUntil;
        }

        private void HandleControllerTick(object? sender, EventArgs e)
        {
            if (!IsActive || !XInputController.TryReadNavigationState(out ControllerNavigationState state))
            {
                ResetControllerNavigation();
                return;
            }

            if (!_nativeControllerNavigationEnabled)
            {
                // LaunchBox is mouse-oriented. Observe XInput only so controller
                // buttons mirrored as keyboard input do not activate this dialog.
                RememberControllerMirrorState(state);
                _controllerPreviousState = state;
                _controllerBaselineReady = true;
                _heldControllerDirection = null;
                _nextControllerRepeatAt = 0;
                return;
            }

            if (!_controllerBaselineReady)
            {
                _controllerPreviousState = state;
                _controllerBaselineReady = true;
                return;
            }

            ControllerNavigationCommand[] directions =
            {
                ControllerNavigationCommand.Up,
                ControllerNavigationCommand.Down,
                ControllerNavigationCommand.Left,
                ControllerNavigationCommand.Right
            };
            ControllerNavigationCommand? newDirection = directions.FirstOrDefault(
                direction => state.IsPressed(direction) && !_controllerPreviousState.IsPressed(direction));
            bool hasNewDirection = newDirection.HasValue
                && state.IsPressed(newDirection.Value)
                && !_controllerPreviousState.IsPressed(newDirection.Value);
            long now = Environment.TickCount64;

            if (hasNewDirection)
            {
                DispatchControllerNavigation(newDirection!.Value);
                _heldControllerDirection = newDirection.Value;
                _nextControllerRepeatAt = now + 340;
            }
            else if (_heldControllerDirection.HasValue
                && state.IsPressed(_heldControllerDirection.Value))
            {
                if (now >= _nextControllerRepeatAt)
                {
                    DispatchControllerNavigation(_heldControllerDirection.Value);
                    _nextControllerRepeatAt = now + 90;
                }
            }
            else
            {
                _heldControllerDirection = null;
                _nextControllerRepeatAt = 0;
            }

            if (state.Activate
                && !_controllerPreviousState.Activate)
            {
                DispatchControllerNavigation(ControllerNavigationCommand.Activate);
            }
            if (state.Cancel && !_controllerPreviousState.Cancel)
            {
                DispatchControllerNavigation(ControllerNavigationCommand.Cancel);
            }

            _controllerPreviousState = state;
        }

        private void RememberControllerMirrorState(ControllerNavigationState state)
        {
            long graceUntil = Environment.TickCount64 + 300;
            foreach (ControllerNavigationCommand command in Enum.GetValues<ControllerNavigationCommand>())
            {
                if (state.IsPressed(command))
                {
                    _keyboardMirrorGraceUntil[command] = graceUntil;
                }
            }
        }

        private static bool IsBigBoxHost()
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

        private void ResetControllerNavigation()
        {
            _controllerPreviousState = default;
            _heldControllerDirection = null;
            _nextControllerRepeatAt = 0;
            _controllerBaselineReady = false;
        }

        private ComboBox? GetOpenProfile()
        {
            return new[] { _translatorProfile, _joyToKeyProfile }
                .FirstOrDefault(profile => profile.IsDropDownOpen);
        }

        private void OpenProfileDropDown(ComboBox profile)
        {
            // A ComboBox popup lives in a separate WPF focus tree. LaunchBox
            // can also translate the same controller A press into a host-level
            // activation. Keep the dropdown open across that duplicate event.
            _guardedControllerProfile = profile;
            _guardedControllerProfileUntil = Environment.TickCount64 + 500;
            profile.IsDropDownOpen = true;
        }

        private void CloseProfileDropDown(ComboBox profile)
        {
            if (ReferenceEquals(_guardedControllerProfile, profile))
            {
                _guardedControllerProfile = null;
                _guardedControllerProfileUntil = 0;
            }
            profile.IsDropDownOpen = false;
        }

        private void HandleProfileDropDownClosed(object? sender, EventArgs e)
        {
            if (sender is not ComboBox profile
                || !ReferenceEquals(profile, _guardedControllerProfile))
            {
                return;
            }

            if (Environment.TickCount64 > _guardedControllerProfileUntil)
            {
                _guardedControllerProfile = null;
                _guardedControllerProfileUntil = 0;
                return;
            }

            // Reopen after the current routed input finishes. This prevents a
            // duplicate activation from producing a visible open/close flicker.
            Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
            {
                if (ReferenceEquals(profile, _guardedControllerProfile)
                    && Environment.TickCount64 <= _guardedControllerProfileUntil)
                {
                    profile.IsDropDownOpen = true;
                }
            }));
        }

        private void DispatchControllerNavigation(ControllerNavigationCommand command)
        {
            _lastNativeNavigationAt[command] = Environment.TickCount64;

            ComboBox? openProfile = GetOpenProfile();
            if (openProfile != null)
            {
                switch (command)
                {
                    case ControllerNavigationCommand.Up:
                        MoveOpenProfileSelection(openProfile, -1);
                        break;
                    case ControllerNavigationCommand.Down:
                        MoveOpenProfileSelection(openProfile, 1);
                        break;
                    case ControllerNavigationCommand.Activate:
                    case ControllerNavigationCommand.Cancel:
                        CloseProfileDropDown(openProfile);
                        break;
                }
                return;
            }

            switch (command)
            {
                case ControllerNavigationCommand.Up:
                    MoveFocusVertical(-1);
                    break;
                case ControllerNavigationCommand.Down:
                    MoveFocusVertical(1);
                    break;
                case ControllerNavigationCommand.Left:
                    MoveFocusHorizontal(-1);
                    break;
                case ControllerNavigationCommand.Right:
                    MoveFocusHorizontal(1);
                    break;
                case ControllerNavigationCommand.Activate:
                    ActivateFocusedControl();
                    break;
                case ControllerNavigationCommand.Cancel:
                    _cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    break;
            }
        }

        private static void MoveOpenProfileSelection(ComboBox profile, int direction)
        {
            if (profile.Items.Count == 0)
            {
                return;
            }

            int current = profile.SelectedIndex < 0 ? 0 : profile.SelectedIndex;
            profile.SelectedIndex = Math.Clamp(current + direction, 0, profile.Items.Count - 1);
        }

        private bool ActivateFocusedControl()
        {
            if (_translatorEnabled.IsKeyboardFocused)
            {
                _translatorEnabled.IsChecked = _translatorEnabled.IsChecked != true;
                return true;
            }
            else if (_translatorProfile.IsKeyboardFocusWithin)
            {
                OpenProfileDropDown(_translatorProfile);
                return true;
            }
            else if (_refreshTranslatorProfiles.IsKeyboardFocused)
            {
                _refreshTranslatorProfiles.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_joyToKeyEnabled.IsKeyboardFocused)
            {
                _joyToKeyEnabled.IsChecked = _joyToKeyEnabled.IsChecked != true;
                return true;
            }
            else if (_joyToKeyProfile.IsKeyboardFocusWithin)
            {
                OpenProfileDropDown(_joyToKeyProfile);
                return true;
            }
            else if (_refreshJoyToKeyProfiles.IsKeyboardFocused)
            {
                _refreshJoyToKeyProfiles.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_browseTranslator.IsKeyboardFocused)
            {
                _browseTranslator.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_browseJoyToKey.IsKeyboardFocused)
            {
                _browseJoyToKey.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_browseProfiles.IsKeyboardFocused)
            {
                _browseProfiles.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_save.IsKeyboardFocused)
            {
                _save.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            else if (_cancel.IsKeyboardFocused)
            {
                _cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                return true;
            }
            return false;
        }

        private void MoveFocusVertical(int direction)
        {
            Control? current = _focusRows
                .SelectMany(row => row)
                .FirstOrDefault(control => control.IsKeyboardFocusWithin);
            if (current == null)
            {
                FocusFirstAvailableControl();
                return;
            }

            int currentRow = _focusRows.FindIndex(row => row.Contains(current));
            double currentCenter = GetHorizontalCenter(current);
            for (int rowIndex = currentRow + direction;
                rowIndex >= 0 && rowIndex < _focusRows.Count;
                rowIndex += direction)
            {
                Control[] candidates = _focusRows[rowIndex]
                    .Where(IsAvailableForNavigation)
                    .ToArray();
                if (candidates.Length > 0)
                {
                    Control candidate = candidates
                        .OrderBy(control => Math.Abs(GetHorizontalCenter(control) - currentCenter))
                        .First();
                    candidate.Focus();
                    return;
                }
            }
        }

        private void MoveFocusHorizontal(int direction)
        {
            Control[]? currentRow = _focusRows.FirstOrDefault(
                row => row.Any(control => control.IsKeyboardFocusWithin));
            if (currentRow == null)
            {
                FocusFirstAvailableControl();
                return;
            }

            Control[] candidates = currentRow
                .Where(IsAvailableForNavigation)
                .ToArray();
            int current = Array.FindIndex(
                candidates,
                control => control.IsKeyboardFocusWithin);
            int next = current + direction;
            if (current >= 0 && next >= 0 && next < candidates.Length)
            {
                candidates[next].Focus();
            }
        }

        private void FocusFirstAvailableControl()
        {
            Control? first = _focusRows
                .SelectMany(row => row)
                .FirstOrDefault(IsAvailableForNavigation);
            first?.Focus();
        }

        private static bool IsAvailableForNavigation(Control control)
        {
            return control.IsEnabled
                && control.Visibility == Visibility.Visible
                && control.Focusable;
        }

        private double GetHorizontalCenter(Control control)
        {
            try
            {
                Point origin = control.TranslatePoint(new Point(0, 0), this);
                return origin.X + (control.ActualWidth / 2.0);
            }
            catch (InvalidOperationException)
            {
                return 0;
            }
        }

        private static Grid MakeProfileRow(out ComboBox profile, out Button refresh)
        {
            Grid row = new Grid { Margin = new Thickness(54, 0, 18, 10) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBlock label = new TextBlock
            {
                Text = "Profile:",
                Foreground = PrimaryForeground,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 14, 0)
            };
            row.Children.Add(label);

            profile = new ComboBox
            {
                MinHeight = 40,
                Background = ControlBackground,
                Foreground = PrimaryForeground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 5, 10, 5),
                VerticalContentAlignment = VerticalAlignment.Center,
                Style = MakeProfileComboBoxStyle()
            };
            Grid.SetColumn(profile, 1);
            row.Children.Add(profile);

            refresh = MakeButton("Refresh", 116);
            refresh.Margin = new Thickness(12, 0, 0, 0);
            Grid.SetColumn(refresh, 2);
            row.Children.Add(refresh);
            return row;
        }

        private static CheckBox MakeCheckBox(string text, bool isChecked)
        {
            return new CheckBox
            {
                Content = text,
                IsChecked = isChecked,
                Foreground = PrimaryForeground,
                FontSize = 19,
                VerticalContentAlignment = VerticalAlignment.Center,
                Style = MakeCheckBoxStyle()
            };
        }

        private static Button MakeButton(string text, double width)
        {
            Button button = new Button
            {
                Content = text,
                Width = width,
                MinHeight = 46,
                Padding = new Thickness(16, 7, 16, 7),
                Background = ControlBackground,
                Foreground = PrimaryForeground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                FontSize = 18,
                Style = MakeButtonStyle()
            };
            return button;
        }

        private static Style MakeProfileComboBoxStyle()
        {
            // LaunchBox and Big Box expose different application-level ComboBox
            // themes. Keep this selector self-contained so its field and popup
            // retain the same contrast in either host.
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type ComboBox}"">
  <Setter Property=""Background"" Value=""#333438"" />
  <Setter Property=""Foreground"" Value=""#F3F4F6"" />
  <Setter Property=""BorderBrush"" Value=""#686B72"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Padding"" Value=""10,5,10,5"" />
  <Setter Property=""MaxDropDownHeight"" Value=""280"" />
  <Setter Property=""ScrollViewer.CanContentScroll"" Value=""True"" />
  <Setter Property=""ItemContainerStyle"">
    <Setter.Value>
      <Style TargetType=""{x:Type ComboBoxItem}"">
        <Setter Property=""Background"" Value=""#333438"" />
        <Setter Property=""Foreground"" Value=""#F3F4F6"" />
        <Setter Property=""Padding"" Value=""10,7"" />
        <Setter Property=""HorizontalContentAlignment"" Value=""Stretch"" />
        <Setter Property=""Template"">
          <Setter.Value>
            <ControlTemplate TargetType=""{x:Type ComboBoxItem}"">
              <Border x:Name=""ItemBorder""
                      Background=""{TemplateBinding Background}""
                      Padding=""{TemplateBinding Padding}""
                      SnapsToDevicePixels=""True"">
                <ContentPresenter VerticalAlignment=""Center"" />
              </Border>
              <ControlTemplate.Triggers>
                <Trigger Property=""IsHighlighted"" Value=""True"">
                  <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#1683D8"" />
                  <Setter Property=""Foreground"" Value=""#FFFFFF"" />
                </Trigger>
                <Trigger Property=""IsSelected"" Value=""True"">
                  <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#126CB2"" />
                  <Setter Property=""Foreground"" Value=""#FFFFFF"" />
                </Trigger>
                <Trigger Property=""IsEnabled"" Value=""False"">
                  <Setter Property=""Opacity"" Value=""0.48"" />
                </Trigger>
              </ControlTemplate.Triggers>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Style>
    </Setter.Value>
  </Setter>
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type ComboBox}"">
        <Grid SnapsToDevicePixels=""True"">
          <Border x:Name=""FieldBorder""
                  Background=""#333438""
                  BorderBrush=""#686B72""
                  BorderThickness=""1"">
            <Grid>
              <TextBlock x:Name=""SelectionText""
                         Text=""{TemplateBinding SelectionBoxItem}""
                         Foreground=""#F3F4F6""
                         Margin=""10,5,38,5""
                         VerticalAlignment=""Center""
                         TextTrimming=""CharacterEllipsis"" />
              <Path Data=""M 0 0 L 5 5 L 10 0 Z""
                    Fill=""#F3F4F6""
                    HorizontalAlignment=""Right""
                    VerticalAlignment=""Center""
                    Margin=""0,0,13,0"" />
            </Grid>
          </Border>
          <ToggleButton Focusable=""False""
                        ClickMode=""Press""
                        Background=""Transparent""
                        BorderThickness=""0""
                        IsChecked=""{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}"">
            <ToggleButton.Template>
              <ControlTemplate TargetType=""{x:Type ToggleButton}"">
                <Border Background=""Transparent"" />
              </ControlTemplate>
            </ToggleButton.Template>
          </ToggleButton>
          <Popup x:Name=""PART_Popup""
                 Placement=""Bottom""
                 AllowsTransparency=""True""
                 Focusable=""False""
                 IsOpen=""{TemplateBinding IsDropDownOpen}""
                 PopupAnimation=""Fade"">
            <Border Background=""#333438""
                    BorderBrush=""#686B72""
                    BorderThickness=""1""
                    MinWidth=""{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}""
                    MaxHeight=""{TemplateBinding MaxDropDownHeight}"">
              <ScrollViewer CanContentScroll=""True"">
                <ItemsPresenter KeyboardNavigation.DirectionalNavigation=""Contained"" />
              </ScrollViewer>
            </Border>
          </Popup>
        </Grid>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsKeyboardFocusWithin"" Value=""True"">
            <Setter TargetName=""FieldBorder"" Property=""BorderBrush"" Value=""#1683D8"" />
            <Setter TargetName=""FieldBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""FieldBorder"" Property=""BorderBrush"" Value=""#1683D8"" />
          </Trigger>
          <Trigger Property=""IsEnabled"" Value=""False"">
            <Setter TargetName=""FieldBorder"" Property=""Opacity"" Value=""0.48"" />
            <Setter TargetName=""SelectionText"" Property=""Foreground"" Value=""#AEB3BC"" />
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>";

            return (Style)XamlReader.Parse(xaml);
        }

        private static Style MakeButtonStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type Button}"">
  <Setter Property=""Background"" Value=""#333438"" />
  <Setter Property=""Foreground"" Value=""#F3F4F6"" />
  <Setter Property=""BorderBrush"" Value=""#686B72"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type Button}"">
        <Border x:Name=""ButtonBorder""
                Background=""{TemplateBinding Background}""
                BorderBrush=""{TemplateBinding BorderBrush}""
                BorderThickness=""{TemplateBinding BorderThickness}""
                Padding=""{TemplateBinding Padding}""
                SnapsToDevicePixels=""True"">
          <ContentPresenter HorizontalAlignment=""Center""
                            VerticalAlignment=""Center""
                            RecognizesAccessKey=""True"" />
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#3D3F44"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#8A8E96"" />
          </Trigger>
          <Trigger Property=""IsKeyboardFocused"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#234A68"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#4FB3FF"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
          <Trigger Property=""IsPressed"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#126CB2"" />
          </Trigger>
          <Trigger Property=""IsEnabled"" Value=""False"">
            <Setter TargetName=""ButtonBorder"" Property=""Opacity"" Value=""0.48"" />
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>";

            return (Style)XamlReader.Parse(xaml);
        }

        private static Style MakeCheckBoxStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type CheckBox}"">
  <Setter Property=""Foreground"" Value=""#F3F4F6"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type CheckBox}"">
        <Border x:Name=""FocusBorder""
                Background=""Transparent""
                BorderBrush=""Transparent""
                BorderThickness=""2""
                Padding=""6,4""
                SnapsToDevicePixels=""True"">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width=""Auto"" />
              <ColumnDefinition Width=""*"" />
            </Grid.ColumnDefinitions>
            <Border x:Name=""CheckBoxBorder""
                    Width=""18""
                    Height=""18""
                    Background=""#333438""
                    BorderBrush=""#858991""
                    BorderThickness=""1""
                    VerticalAlignment=""Center"">
              <Path x:Name=""CheckMark""
                    Data=""M 3 8 L 7 12 L 15 3""
                    Stroke=""#FFFFFF""
                    StrokeThickness=""2""
                    StrokeStartLineCap=""Round""
                    StrokeEndLineCap=""Round""
                    Visibility=""Collapsed"" />
            </Border>
            <ContentPresenter Grid.Column=""1""
                              Margin=""10,0,0,0""
                              VerticalAlignment=""Center""
                              RecognizesAccessKey=""True"" />
          </Grid>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsChecked"" Value=""True"">
            <Setter TargetName=""CheckBoxBorder"" Property=""Background"" Value=""#1683D8"" />
            <Setter TargetName=""CheckBoxBorder"" Property=""BorderBrush"" Value=""#4FB3FF"" />
            <Setter TargetName=""CheckMark"" Property=""Visibility"" Value=""Visible"" />
          </Trigger>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""FocusBorder"" Property=""Background"" Value=""#303238"" />
          </Trigger>
          <Trigger Property=""IsKeyboardFocused"" Value=""True"">
            <Setter TargetName=""FocusBorder"" Property=""Background"" Value=""#233748"" />
            <Setter TargetName=""FocusBorder"" Property=""BorderBrush"" Value=""#4FB3FF"" />
          </Trigger>
          <Trigger Property=""IsEnabled"" Value=""False"">
            <Setter TargetName=""FocusBorder"" Property=""Opacity"" Value=""0.48"" />
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>";

            return (Style)XamlReader.Parse(xaml);
        }

        private static SolidColorBrush BrushFrom(string value)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
        }
    }
}
