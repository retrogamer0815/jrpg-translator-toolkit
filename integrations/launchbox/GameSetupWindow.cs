using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using Ellipse = System.Windows.Shapes.Ellipse;

// Companion source snapshot: JRPG Translator v0.9.5.

namespace JrpgTranslator.LaunchBox
{
    public sealed class GameSetupWindow : Window
    {
        private static readonly Brush WindowBackground = BrushFrom("#202127");
        private static readonly Brush PanelBackground = BrushFrom("#292C34");
        private static readonly Brush ControlBackground = BrushFrom("#242730");
        private static readonly Brush PrimaryForeground = BrushFrom("#ECEEF3");
        private static readonly Brush MutedForeground = BrushFrom("#AEB8C9");
        private static readonly Brush ControlBorderBrush = BrushFrom("#4B5260");
        private static readonly Brush AccentBrush = BrushFrom("#168ED1");
        private static readonly Brush ReadyBrush = BrushFrom("#54C88A");
        private static readonly Brush WarningBrush = BrushFrom("#F2B45B");
        private static readonly object CurrentSettingsProfile = new CurrentSettingsProfileChoice();

        // A distinct item keeps the display label from colliding with a real
        // Profile name. Only real string items are persisted as Profile names.
        private sealed class CurrentSettingsProfileChoice
        {
            public override string ToString() => "None — use current settings";
        }

        private readonly PluginConfiguration _configuration;
        private readonly CheckBox _translatorEnabled;
        private readonly CheckBox _joyToKeyEnabled;
        private readonly ComboBox _translatorProfile;
        private readonly Button _refreshTranslatorProfiles;
        private readonly Button _openTranslator;
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
        private readonly TextBlock _translatorReadiness;
        private readonly TextBlock _joyToKeyReadiness;
        private readonly Ellipse _translatorReadinessDot;
        private readonly Ellipse _joyToKeyReadinessDot;
        private readonly TextBlock _translatorPathValue;
        private readonly TextBlock _joyToKeyPathValue;
        private readonly TextBlock _joyToKeyProfilesPathValue;
        private readonly StackPanel _translatorSettings;
        private readonly StackPanel _joyToKeySettings;
        private readonly StackPanel _locationsPanel;
        private readonly Button _locationsToggle;
        private readonly Button _detectAgain;
        private readonly ScrollViewer _contentScroller;
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
        private bool _locationsExpanded;

        public GameConfiguration Result { get; private set; }

        public GameSetupWindow(PluginConfiguration configuration, GameConfiguration game)
        {
            _configuration = configuration;
            Result = game;
            _nativeControllerNavigationEnabled = IsBigBoxHost();

            Title = "JRPG Translator Setup";
            Rect workArea = SystemParameters.WorkArea;
            double availableWidth = Math.Max(1, workArea.Width - 24);
            double availableHeight = Math.Max(1, workArea.Height - 24);
            Width = Math.Min(920, availableWidth);
            Height = Math.Min(800, availableHeight);
            MinWidth = Math.Min(760, availableWidth);
            MinHeight = Math.Min(560, availableHeight);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanResize;
            ShowInTaskbar = false;
            Background = WindowBackground;
            Foreground = PrimaryForeground;
            Cursor = Cursors.Arrow;
            ForceCursor = true;
            FontFamily = new FontFamily("Segoe UI");
            FontSize = 16;

            Grid root = new Grid { Margin = new Thickness(30, 24, 30, 22) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            TextBlock heading = new TextBlock
            {
                Text = game.GameTitle,
                FontSize = 28,
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
                Margin = new Thickness(0, 0, 0, 20)
            };
            Grid.SetRow(introduction, 1);
            root.Children.Add(introduction);

            StackPanel scrollingContent = new StackPanel();
            _contentScroller = new ScrollViewer
            {
                Content = scrollingContent,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                // The content is composed of a few large panels. Logical scrolling
                // would jump from one whole panel to the next and skip the controls
                // between them, so use physical (pixel-based) scrolling here.
                CanContentScroll = false,
                Focusable = false
            };
            Grid.SetRow(_contentScroller, 2);
            root.Children.Add(_contentScroller);

            Border translatorCard = MakeCard();
            StackPanel translatorCardContent = new StackPanel();
            translatorCard.Child = translatorCardContent;
            scrollingContent.Children.Add(translatorCard);

            _translatorReadinessDot = MakeStatusDot();
            _translatorReadiness = MakeStatusText();
            _translatorEnabled = MakeToggle(game.TranslatorEnabled);
            _translatorEnabled.Checked += (_, _) => UpdateTranslatorControls();
            _translatorEnabled.Unchecked += (_, _) => UpdateTranslatorControls();
            translatorCardContent.Children.Add(MakeIntegrationHeader(
                MakeTranslatorIcon(),
                "JRPG Translator",
                "Capture and translate game text automatically.",
                _translatorReadinessDot,
                _translatorReadiness,
                _translatorEnabled));

            _translatorSettings = new StackPanel
            {
                Margin = new Thickness(96, 0, 24, 22)
            };
            translatorCardContent.Children.Add(_translatorSettings);
            _translatorSettings.Children.Add(MakeProfileSection(
                out _translatorProfile,
                out _translatorProfileStatus));
            _translatorProfile.DropDownClosed += HandleProfileDropDownClosed;
            _translatorProfile.DropDownOpened += (_, _) =>
                RefreshTranslatorProfiles(SelectedTranslatorProfile());
            _refreshTranslatorProfiles = MakeButton("Refresh", 116);
            _refreshTranslatorProfiles.Visibility = Visibility.Collapsed;
            _refreshTranslatorProfiles.Click += (_, _) => RefreshTranslatorProfiles(SelectedTranslatorProfile());

            _openTranslator = MakeButton("Open JRPG Translator…", 224);
            _openTranslator.HorizontalAlignment = HorizontalAlignment.Left;
            _openTranslator.Margin = new Thickness(0, 14, 0, 10);
            _openTranslator.Click += OpenTranslatorClicked;
            _translatorSettings.Children.Add(_openTranslator);

            TextBlock openTranslatorHelp = new TextBlock
            {
                Text = "Configure API keys, translation models, and controller shortcuts in the main app.",
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap
            };
            _translatorSettings.Children.Add(openTranslatorHelp);

            Border joyToKeyCard = MakeCard();
            StackPanel joyToKeyCardContent = new StackPanel();
            joyToKeyCard.Child = joyToKeyCardContent;
            scrollingContent.Children.Add(joyToKeyCard);

            _joyToKeyReadinessDot = MakeStatusDot();
            _joyToKeyReadiness = MakeStatusText();
            _joyToKeyEnabled = MakeToggle(game.JoyToKeyEnabled);
            _joyToKeyEnabled.Checked += (_, _) => UpdateJoyToKeyControls();
            _joyToKeyEnabled.Unchecked += (_, _) => UpdateJoyToKeyControls();
            joyToKeyCardContent.Children.Add(MakeIntegrationHeader(
                MakeJoyToKeyIcon(),
                "JoyToKey",
                "Apply a controller mapping when this game starts.",
                _joyToKeyReadinessDot,
                _joyToKeyReadiness,
                _joyToKeyEnabled));

            _joyToKeySettings = new StackPanel
            {
                Margin = new Thickness(96, 0, 24, 22)
            };
            joyToKeyCardContent.Children.Add(_joyToKeySettings);
            _joyToKeySettings.Children.Add(MakeProfileSection(
                out _joyToKeyProfile,
                out _joyToKeyProfileStatus));
            _joyToKeyProfile.DropDownClosed += HandleProfileDropDownClosed;
            _joyToKeyProfile.DropDownOpened += (_, _) =>
                RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);
            _refreshJoyToKeyProfiles = MakeButton("Refresh", 116);
            _refreshJoyToKeyProfiles.Visibility = Visibility.Collapsed;
            _refreshJoyToKeyProfiles.Click += (_, _) => RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);

            Border locationsCard = MakeCard();
            locationsCard.Margin = new Thickness(0, 0, 0, 4);
            StackPanel locationsCardContent = new StackPanel();
            locationsCard.Child = locationsCardContent;
            scrollingContent.Children.Add(locationsCard);

            Grid locationsHeader = new Grid { Margin = new Thickness(18, 12, 18, 12) };
            locationsHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            locationsHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            locationsCardContent.Children.Add(locationsHeader);

            _locationsToggle = MakeDisclosureButton("Application locations");
            _locationsToggle.Click += (_, _) => SetLocationsExpanded(!_locationsExpanded);
            locationsHeader.Children.Add(_locationsToggle);

            _readiness = new TextBlock
            {
                Foreground = MutedForeground,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Right,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(_readiness, 1);
            locationsHeader.Children.Add(_readiness);

            _locationsPanel = new StackPanel
            {
                Margin = new Thickness(24, 0, 24, 22),
                Visibility = Visibility.Collapsed
            };
            locationsCardContent.Children.Add(_locationsPanel);
            _locationsPanel.Children.Add(new TextBlock
            {
                Text = "Locations are detected automatically. Choose a different file or folder only when detection needs help.",
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14)
            });

            _browseTranslator = MakeButton("Browse…", 108);
            _browseTranslator.Click += BrowseTranslatorClicked;
            _translatorPathValue = MakePathValue();
            _locationsPanel.Children.Add(MakeLocationRow(
                "JRPG Translator executable", _translatorPathValue, _browseTranslator));

            _browseJoyToKey = MakeButton("Browse…", 108);
            _browseJoyToKey.Click += BrowseJoyToKeyClicked;
            _joyToKeyPathValue = MakePathValue();
            _locationsPanel.Children.Add(MakeLocationRow(
                "JoyToKey executable", _joyToKeyPathValue, _browseJoyToKey));

            _browseProfiles = MakeButton("Browse…", 108);
            _browseProfiles.Click += BrowseProfilesClicked;
            _joyToKeyProfilesPathValue = MakePathValue();
            _locationsPanel.Children.Add(MakeLocationRow(
                "JoyToKey profiles folder", _joyToKeyProfilesPathValue, _browseProfiles));

            Grid locationActions = new Grid { Margin = new Thickness(0, 4, 0, 0) };
            locationActions.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            locationActions.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _locationsPanel.Children.Add(locationActions);

            TextBlock stageNotice = new TextBlock
            {
                Text = "Plugin-started tools are closed or restored automatically after the game exits.",
                Foreground = MutedForeground,
                FontStyle = FontStyles.Italic,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 16, 0)
            };
            locationActions.Children.Add(stageNotice);

            _detectAgain = MakeButton("Detect again", 132);
            _detectAgain.Click += DetectAgainClicked;
            Grid.SetColumn(_detectAgain, 1);
            locationActions.Children.Add(_detectAgain);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 18, 0, 0)
            };
            Grid.SetRow(buttons, 3);
            root.Children.Add(buttons);

            _cancel = MakeButton("Cancel", 150);
            _cancel.IsCancel = true;
            _cancel.Click += (_, _) => { DialogResult = false; };
            buttons.Children.Add(_cancel);

            _save = MakeButton("Save", 150);
            _save.IsDefault = true;
            _save.Margin = new Thickness(12, 0, 0, 0);
            _save.Background = AccentBrush;
            _save.BorderBrush = AccentBrush;
            _save.Style = MakePrimaryButtonStyle();
            _save.Click += SaveClicked;
            buttons.Children.Add(_save);

            _focusRows = new List<Control[]>
            {
                new Control[] { _translatorEnabled },
                new Control[] { _translatorProfile },
                new Control[] { _openTranslator },
                new Control[] { _joyToKeyEnabled },
                new Control[] { _joyToKeyProfile },
                new Control[] { _locationsToggle },
                new Control[] { _browseTranslator },
                new Control[] { _browseJoyToKey },
                new Control[] { _browseProfiles },
                new Control[] { _detectAgain },
                new Control[] { _cancel, _save }
            };

            foreach (Control control in _focusRows.SelectMany(row => row))
            {
                control.GotKeyboardFocus += BringFocusedControlIntoView;
            }

            _controllerTimer = new DispatcherTimer(DispatcherPriority.Input)
            {
                Interval = TimeSpan.FromMilliseconds(20)
            };
            _controllerTimer.Tick += HandleControllerTick;

            PreviewKeyDown += HandlePreviewKeyDown;
            PreviewKeyUp += HandlePreviewKeyUp;
            SourceInitialized += (_, _) => ApplyDarkTitleBar();
            Loaded += (_, _) =>
            {
                FitWindowToVisibleWorkArea();
                // LaunchBox can keep its busy cursor active while a synchronous
                // plugin command is open. This dialog has no background work.
                Mouse.OverrideCursor = null;
                Cursor = Cursors.Arrow;
                ForceCursor = true;
                _translatorEnabled.Focus();
                _controllerTimer.Start();
            };
            Activated += (_, _) =>
            {
                RefreshTranslatorProfiles(SelectedTranslatorProfile());
                RefreshJoyToKeyProfiles(_joyToKeyProfile.Text);
                UpdatePathStatus();
                ResetControllerNavigation();
            };
            Deactivated += (_, _) => ResetControllerNavigation();
            Closed += HandleClosed;

            RefreshTranslatorProfiles(game.TranslatorProfile);
            RefreshJoyToKeyProfiles(game.JoyToKeyProfile);
            UpdateTranslatorControls();
            UpdateJoyToKeyControls();
            UpdatePathStatus();
        }

        private void BringFocusedControlIntoView(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!IsLoaded || sender is not FrameworkElement element)
            {
                return;
            }

            Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(() =>
            {
                if (IsLoaded && element.IsLoaded)
                {
                    element.BringIntoView();
                }
            }));
        }

        private void HandleClosed(object? sender, EventArgs e)
        {
            _controllerTimer.Stop();
            _controllerTimer.Tick -= HandleControllerTick;
            ResetControllerNavigation();
        }

        private void FitWindowToVisibleWorkArea()
        {
            IntPtr windowHandle = new WindowInteropHelper(this).Handle;
            if (windowHandle == IntPtr.Zero)
            {
                return;
            }

            IntPtr referenceHandle = Owner == null
                ? windowHandle
                : new WindowInteropHelper(Owner).Handle;
            IntPtr monitor = MonitorFromWindow(
                referenceHandle == IntPtr.Zero ? windowHandle : referenceHandle,
                MonitorDefaultToNearest);
            MonitorInfo monitorInfo = new MonitorInfo
            {
                Size = Marshal.SizeOf<MonitorInfo>()
            };
            if (monitor == IntPtr.Zero || !GetMonitorInfo(monitor, ref monitorInfo))
            {
                return;
            }

            DpiScale dpi = VisualTreeHelper.GetDpi(this);
            double scaleX = Math.Max(1, dpi.DpiScaleX);
            double scaleY = Math.Max(1, dpi.DpiScaleY);
            const double logicalMargin = 12;
            double availableWidth = Math.Max(
                1,
                (monitorInfo.WorkArea.Right - monitorInfo.WorkArea.Left) / scaleX
                    - logicalMargin * 2);
            double availableHeight = Math.Max(
                1,
                (monitorInfo.WorkArea.Bottom - monitorInfo.WorkArea.Top) / scaleY
                    - logicalMargin * 2);

            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            MinWidth = Math.Min(720, availableWidth);
            MinHeight = Math.Min(500, availableHeight);

            double targetWidth = Math.Min(
                double.IsNaN(Width) || Width <= 0 ? 840 : Width,
                availableWidth);
            double targetHeight = Math.Min(
                double.IsNaN(Height) || Height <= 0 ? 900 : Height,
                availableHeight);
            Width = targetWidth;
            Height = targetHeight;

            int physicalWidth = Math.Max(1, (int)Math.Round(targetWidth * scaleX));
            int physicalHeight = Math.Max(1, (int)Math.Round(targetHeight * scaleY));
            int centerX = (monitorInfo.WorkArea.Left + monitorInfo.WorkArea.Right) / 2;
            int centerY = (monitorInfo.WorkArea.Top + monitorInfo.WorkArea.Bottom) / 2;
            if (Owner != null
                && referenceHandle != IntPtr.Zero
                && GetWindowRect(referenceHandle, out NativeRect ownerBounds))
            {
                centerX = (ownerBounds.Left + ownerBounds.Right) / 2;
                centerY = (ownerBounds.Top + ownerBounds.Bottom) / 2;
            }

            int marginX = Math.Max(0, (int)Math.Round(logicalMargin * scaleX));
            int marginY = Math.Max(0, (int)Math.Round(logicalMargin * scaleY));
            int left = centerX - physicalWidth / 2;
            int top = centerY - physicalHeight / 2;
            left = Math.Max(
                monitorInfo.WorkArea.Left + marginX,
                Math.Min(left, monitorInfo.WorkArea.Right - marginX - physicalWidth));
            top = Math.Max(
                monitorInfo.WorkArea.Top + marginY,
                Math.Min(top, monitorInfo.WorkArea.Bottom - marginY - physicalHeight));

            SetWindowPos(
                windowHandle,
                IntPtr.Zero,
                left,
                top,
                physicalWidth,
                physicalHeight,
                SetWindowPositionNoActivate | SetWindowPositionNoZOrder);
        }

        private const uint MonitorDefaultToNearest = 2;
        private const uint SetWindowPositionNoZOrder = 0x0004;
        private const uint SetWindowPositionNoActivate = 0x0010;

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MonitorInfo
        {
            public int Size;
            public NativeRect MonitorArea;
            public NativeRect WorkArea;
            public uint Flags;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr windowHandle, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo monitorInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr windowHandle, out NativeRect bounds);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(
            IntPtr windowHandle,
            IntPtr insertAfter,
            int left,
            int top,
            int width,
            int height,
            uint flags);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr windowHandle,
            int attribute,
            ref int value,
            int valueSize);

        private void ApplyDarkTitleBar()
        {
            IntPtr windowHandle = new WindowInteropHelper(this).Handle;
            if (windowHandle == IntPtr.Zero)
            {
                return;
            }

            int enabled = 1;
            const int immersiveDarkMode = 20;
            if (DwmSetWindowAttribute(
                    windowHandle,
                    immersiveDarkMode,
                    ref enabled,
                    sizeof(int)) != 0)
            {
                const int immersiveDarkModeBefore20H1 = 19;
                DwmSetWindowAttribute(
                    windowHandle,
                    immersiveDarkModeBefore20H1,
                    ref enabled,
                    sizeof(int));
            }
        }

        private void RefreshTranslatorProfiles(string preferredProfile)
        {
            IReadOnlyList<string> profiles = TranslatorProfileDiscovery.GetProfiles(_configuration);
            List<object> choices = new List<object> { CurrentSettingsProfile };
            choices.AddRange(profiles);
            _translatorProfile.ItemsSource = choices;

            string? selected = profiles.FirstOrDefault(
                value => string.Equals(value, preferredProfile, StringComparison.OrdinalIgnoreCase));
            _translatorProfile.SelectedItem = selected is null ? CurrentSettingsProfile : selected;

            string directory = TranslatorProfileDiscovery.GetProfilesDirectory(_configuration);
            _translatorProfileStatus.Text = profiles.Count == 0
                ? "No saved Profiles found"
                : profiles.Count + (profiles.Count == 1 ? " Profile available" : " Profiles available");
            _translatorProfileStatus.ToolTip = profiles.Count == 0
                ? "No Profile files were found in " + DisplayPath(directory)
                : DisplayPath(directory);
        }

        private string SelectedTranslatorProfile()
        {
            return _translatorProfile.SelectedItem as string ?? string.Empty;
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
                ? "No profiles found"
                : profiles.Count + (profiles.Count == 1 ? " profile available" : " profiles available");
            _joyToKeyProfileStatus.ToolTip = DisplayPath(_configuration.JoyToKeyProfilesDirectory);
        }

        private string BuildReadinessText()
        {
            string translator = PluginPaths.ResolveTranslatorExecutable(_configuration);
            bool translatorReady = File.Exists(translator);
            bool joyToKeyReady = File.Exists(_configuration.JoyToKeyExecutable)
                && Directory.Exists(_configuration.JoyToKeyProfilesDirectory);
            if (translatorReady && joyToKeyReady)
            {
                return "Everything ready";
            }

            int missing = (translatorReady ? 0 : 1) + (joyToKeyReady ? 0 : 1);
            return missing == 1 ? "1 item needs setup" : "2 items need setup";
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
            _joyToKeyEnabled.Content = enabled ? "On" : "Off";
            _joyToKeySettings.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateTranslatorControls()
        {
            bool enabled = _translatorEnabled.IsChecked == true;
            _translatorProfile.IsEnabled = enabled;
            _refreshTranslatorProfiles.IsEnabled = enabled;
            _translatorEnabled.Content = enabled ? "On" : "Off";
            _translatorSettings.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;

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
            RefreshTranslatorProfiles(SelectedTranslatorProfile());
            UpdatePathStatus();
        }

        private void OpenTranslatorClicked(object sender, RoutedEventArgs e)
        {
            string executable = PluginPaths.ResolveTranslatorExecutable(_configuration);
            if (!File.Exists(executable))
            {
                MessageBox.Show(
                    this,
                    "JRPG Translator.exe was not found. Open Application locations and select it with Browse.",
                    "JRPG Translator Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = executable,
                    WorkingDirectory = Path.GetDirectoryName(executable) ?? string.Empty,
                    UseShellExecute = true
                });
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    this,
                    "JRPG Translator could not be opened:\n\n" + exception.Message,
                    "JRPG Translator Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
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
            string translatorPath = PluginPaths.ResolveTranslatorExecutable(_configuration);
            bool translatorReady = File.Exists(translatorPath);
            bool joyToKeyExecutableReady = File.Exists(_configuration.JoyToKeyExecutable);
            bool joyToKeyProfilesReady = Directory.Exists(_configuration.JoyToKeyProfilesDirectory);
            bool joyToKeyReady = joyToKeyExecutableReady && joyToKeyProfilesReady;

            SetReadinessBadge(
                _translatorReadiness,
                _translatorReadinessDot,
                translatorReady,
                translatorReady ? "Ready" : "Needs setup");
            SetReadinessBadge(
                _joyToKeyReadiness,
                _joyToKeyReadinessDot,
                joyToKeyReady,
                joyToKeyReady ? "Ready" : "Needs setup");

            SetPathValue(_translatorPathValue, translatorPath, translatorReady);
            SetPathValue(
                _joyToKeyPathValue,
                _configuration.JoyToKeyExecutable,
                joyToKeyExecutableReady);
            SetPathValue(
                _joyToKeyProfilesPathValue,
                _configuration.JoyToKeyProfilesDirectory,
                joyToKeyProfilesReady);
            _readiness.Text = BuildReadinessText();
            _readiness.Foreground = translatorReady && joyToKeyReady
                ? ReadyBrush
                : WarningBrush;
            if (!translatorReady || !joyToKeyReady)
            {
                SetLocationsExpanded(true);
            }
        }

        private void DetectAgainClicked(object sender, RoutedEventArgs e)
        {
            string translatorProfile = SelectedTranslatorProfile();
            string joyToKeyProfile = _joyToKeyProfile.Text;
            _configuration.TranslatorExecutable = string.Empty;
            _configuration.JoyToKeyExecutable = string.Empty;
            _configuration.JoyToKeyProfilesDirectory = string.Empty;
            PluginPaths.PopulateDetectedPaths(_configuration);
            RefreshTranslatorProfiles(translatorProfile);
            RefreshJoyToKeyProfiles(joyToKeyProfile);
            UpdatePathStatus();
        }

        private void SetLocationsExpanded(bool expanded)
        {
            _locationsExpanded = expanded;
            _locationsPanel.Visibility = expanded ? Visibility.Visible : Visibility.Collapsed;
            _locationsToggle.Content = (expanded ? "▾  " : "▸  ") + "Application locations";
        }

        private static void SetReadinessBadge(
            TextBlock label,
            Ellipse dot,
            bool ready,
            string text)
        {
            label.Text = text;
            label.Foreground = ready ? ReadyBrush : WarningBrush;
            dot.Fill = ready ? ReadyBrush : WarningBrush;
        }

        private static void SetPathValue(TextBlock label, string path, bool ready)
        {
            label.Text = string.IsNullOrWhiteSpace(path) ? "Not detected" : path;
            label.Foreground = ready ? PrimaryForeground : WarningBrush;
            label.ToolTip = label.Text;
        }

        private void SaveClicked(object sender, RoutedEventArgs e)
        {
            if (_translatorEnabled.IsChecked == true
                && _translatorProfile.Items.Count > 0
                && _translatorProfile.SelectedItem == null)
            {
                MessageBox.Show(
                    this,
                    "Select a Profile or None to use current settings.",
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

            SaveSelections();
            DialogResult = true;
        }

        private void SaveSelections()
        {
            Result.TranslatorEnabled = _translatorEnabled.IsChecked == true;
            Result.TranslatorProfile = SelectedTranslatorProfile();
            Result.JoyToKeyEnabled = _joyToKeyEnabled.IsChecked == true;
            Result.JoyToKeyProfile = _joyToKeyProfile.SelectedItem as string ?? string.Empty;
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
                if (e.Key == Key.Escape)
                {
                    CloseProfileDropDown(openProfile);
                    e.Handled = true;
                }
                else if (e.Key == Key.Enter || e.Key == Key.Space)
                {
                    // Let the ComboBox process the key itself. WPF keeps keyboard
                    // navigation as a highlighted candidate while the popup is
                    // open and commits it on Enter/Space. Consuming the preview
                    // event here used to close the popup before that commit.
                    ReleaseProfileDropDownGuard(openProfile);
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
            try
            {
                HandleControllerTickCore();
            }
            catch (Exception exception)
            {
                RuntimeLog.Write("Setup-window controller polling was disabled: " + exception.Message);
                _controllerTimer.Stop();
                ResetControllerNavigation();
            }
        }

        private void HandleControllerTickCore()
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
            return PluginHostEnvironment.IsBigBoxHost();
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
            ReleaseProfileDropDownGuard(profile);
            profile.IsDropDownOpen = false;
        }

        private void ReleaseProfileDropDownGuard(ComboBox profile)
        {
            if (!ReferenceEquals(_guardedControllerProfile, profile))
            {
                return;
            }

            _guardedControllerProfile = null;
            _guardedControllerProfileUntil = 0;
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
                if (IsLoaded
                    && profile.IsLoaded
                    && ReferenceEquals(profile, _guardedControllerProfile)
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
            Control? focused = _focusRows
                .SelectMany(row => row)
                .FirstOrDefault(control => control.IsKeyboardFocusWithin);
            return focused != null && ActivateControl(focused);
        }

        private bool ActivateControl(Control control)
        {
            if (ReferenceEquals(control, _translatorEnabled))
            {
                _translatorEnabled.IsChecked = _translatorEnabled.IsChecked != true;
                return true;
            }
            if (ReferenceEquals(control, _translatorProfile))
            {
                OpenProfileDropDown(_translatorProfile);
                return true;
            }
            if (ReferenceEquals(control, _joyToKeyEnabled))
            {
                _joyToKeyEnabled.IsChecked = _joyToKeyEnabled.IsChecked != true;
                return true;
            }
            if (ReferenceEquals(control, _joyToKeyProfile))
            {
                OpenProfileDropDown(_joyToKeyProfile);
                return true;
            }
            if (control is Button button && button.IsEnabled)
            {
                button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
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
                && control.IsVisible
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

        private static Border MakeCard()
        {
            return new Border
            {
                Background = PanelBackground,
                BorderBrush = BrushFrom("#343A46"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Margin = new Thickness(0, 0, 0, 14),
                ClipToBounds = true
            };
        }

        private static Grid MakeIntegrationHeader(
            FrameworkElement icon,
            string title,
            string subtitle,
            Ellipse readinessDot,
            TextBlock readiness,
            CheckBox toggle)
        {
            Grid header = new Grid { Margin = new Thickness(24, 20, 24, 16) };
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            icon.HorizontalAlignment = HorizontalAlignment.Left;
            icon.VerticalAlignment = VerticalAlignment.Center;
            header.Children.Add(icon);

            StackPanel heading = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            heading.Children.Add(new TextBlock
            {
                Text = title,
                Foreground = PrimaryForeground,
                FontSize = 22,
                FontWeight = FontWeights.SemiBold
            });
            heading.Children.Add(new TextBlock
            {
                Text = subtitle,
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 3, 0, 0)
            });
            Grid.SetColumn(heading, 1);
            header.Children.Add(heading);

            StackPanel status = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(18, 0, 18, 0)
            };
            status.Children.Add(readinessDot);
            status.Children.Add(readiness);
            Grid.SetColumn(status, 2);
            header.Children.Add(status);

            Grid.SetColumn(toggle, 3);
            header.Children.Add(toggle);
            return header;
        }

        private static FrameworkElement MakeTranslatorIcon()
        {
            ImageSource? source = LoadEmbeddedIconSource(
                "JrpgTranslator.LaunchBox.Assets.menu-icon.png");
            if (source != null)
            {
                return new Image
                {
                    Source = source,
                    Width = 58,
                    Height = 58,
                    Stretch = Stretch.Uniform
                };
            }

            return MakeMonogramIcon("JRPG");
        }

        private static ImageSource? LoadEmbeddedIconSource(string resourceName)
        {
            try
            {
                using Stream? stream = typeof(GameSetupWindow).Assembly.GetManifestResourceStream(
                    resourceName);
                if (stream == null)
                {
                    return null;
                }

                BitmapImage image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.DecodePixelWidth = 128;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
                return image;
            }
            catch
            {
                return null;
            }
        }

        private static FrameworkElement MakeJoyToKeyIcon()
        {
            ImageSource? source = LoadEmbeddedIconSource(
                "JrpgTranslator.LaunchBox.Assets.joytokey-icon.png");
            if (source != null)
            {
                return new Image
                {
                    Source = source,
                    Width = 58,
                    Height = 58,
                    Stretch = Stretch.Uniform
                };
            }

            return MakeMonogramIcon("JTK");
        }

        private static Border MakeMonogramIcon(string text)
        {
            return new Border
            {
                Width = 54,
                Height = 54,
                Background = ControlBackground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Child = new TextBlock
                {
                    Text = text,
                    Foreground = MutedForeground,
                    FontWeight = FontWeights.SemiBold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
        }

        private static Ellipse MakeStatusDot()
        {
            return new Ellipse
            {
                Width = 8,
                Height = 8,
                Fill = MutedForeground,
                Margin = new Thickness(0, 1, 7, 0)
            };
        }

        private static TextBlock MakeStatusText()
        {
            return new TextBlock
            {
                Text = "Checking…",
                Foreground = MutedForeground,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        private static StackPanel MakeProfileSection(
            out ComboBox profile,
            out TextBlock status)
        {
            StackPanel section = new StackPanel();

            TextBlock label = new TextBlock
            {
                Text = "Profile",
                Foreground = PrimaryForeground,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 7)
            };
            section.Children.Add(label);

            Grid row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            section.Children.Add(row);

            profile = new ComboBox
            {
                MinHeight = 42,
                Background = ControlBackground,
                Foreground = PrimaryForeground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 5, 10, 5),
                VerticalContentAlignment = VerticalAlignment.Center,
                Style = MakeProfileComboBoxStyle()
            };
            row.Children.Add(profile);

            status = new TextBlock
            {
                Foreground = MutedForeground,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(16, 0, 0, 0),
                MinWidth = 138
            };
            Grid.SetColumn(status, 1);
            row.Children.Add(status);
            return section;
        }

        private static CheckBox MakeToggle(bool isChecked)
        {
            return new CheckBox
            {
                Content = isChecked ? "On" : "Off",
                IsChecked = isChecked,
                Foreground = PrimaryForeground,
                FontSize = 16,
                VerticalContentAlignment = VerticalAlignment.Center,
                Style = MakeToggleStyle()
            };
        }

        private static Button MakeDisclosureButton(string text)
        {
            Button button = MakeButton("▸  " + text, 250);
            button.HorizontalAlignment = HorizontalAlignment.Left;
            button.HorizontalContentAlignment = HorizontalAlignment.Left;
            button.Background = Brushes.Transparent;
            button.BorderBrush = Brushes.Transparent;
            button.FontWeight = FontWeights.SemiBold;
            button.Padding = new Thickness(4, 5, 8, 5);
            return button;
        }

        private static TextBlock MakePathValue()
        {
            return new TextBlock
            {
                Foreground = PrimaryForeground,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        private static StackPanel MakeLocationRow(
            string label,
            TextBlock value,
            Button browse)
        {
            StackPanel section = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            section.Children.Add(new TextBlock
            {
                Text = label,
                Foreground = PrimaryForeground,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6)
            });

            Grid row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            section.Children.Add(row);

            Border field = new Border
            {
                Background = ControlBackground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12, 10, 12, 10),
                Child = value
            };
            row.Children.Add(field);

            browse.Margin = new Thickness(10, 0, 0, 0);
            Grid.SetColumn(browse, 1);
            row.Children.Add(browse);
            return section;
        }

        private static Button MakeButton(string text, double width)
        {
            Button button = new Button
            {
                Content = text,
                Width = width,
                MinHeight = 42,
                Padding = new Thickness(16, 7, 16, 7),
                Background = ControlBackground,
                Foreground = PrimaryForeground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                FontSize = 16,
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
  <Setter Property=""Background"" Value=""#242730"" />
  <Setter Property=""Foreground"" Value=""#ECEEF3"" />
  <Setter Property=""BorderBrush"" Value=""#4B5260"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Padding"" Value=""10,5,10,5"" />
  <Setter Property=""MaxDropDownHeight"" Value=""280"" />
  <Setter Property=""ScrollViewer.CanContentScroll"" Value=""True"" />
  <Setter Property=""ItemContainerStyle"">
    <Setter.Value>
      <Style TargetType=""{x:Type ComboBoxItem}"">
        <Setter Property=""Background"" Value=""#242730"" />
        <Setter Property=""Foreground"" Value=""#ECEEF3"" />
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
                  <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#168ED1"" />
                  <Setter Property=""Foreground"" Value=""#FFFFFF"" />
                </Trigger>
                <Trigger Property=""IsSelected"" Value=""True"">
                  <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#1476AD"" />
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
                  Background=""#242730""
                  BorderBrush=""#4B5260""
                  BorderThickness=""1""
                  CornerRadius=""4"">
            <Grid>
              <TextBlock x:Name=""SelectionText""
                         Text=""{TemplateBinding SelectionBoxItem}""
                         Foreground=""#ECEEF3""
                         Margin=""10,5,38,5""
                         VerticalAlignment=""Center""
                         TextTrimming=""CharacterEllipsis"" />
              <Path Data=""M 0 0 L 5 5 L 10 0 Z""
                    Fill=""#ECEEF3""
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
            <Border Background=""#242730""
                    BorderBrush=""#4B5260""
                    BorderThickness=""1""
                    CornerRadius=""4""
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
            <Setter TargetName=""FieldBorder"" Property=""BorderBrush"" Value=""#56C4F5"" />
            <Setter TargetName=""FieldBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""FieldBorder"" Property=""BorderBrush"" Value=""#56C4F5"" />
          </Trigger>
          <Trigger Property=""IsEnabled"" Value=""False"">
            <Setter TargetName=""FieldBorder"" Property=""Opacity"" Value=""0.48"" />
            <Setter TargetName=""SelectionText"" Property=""Foreground"" Value=""#AEB8C9"" />
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
  <Setter Property=""Background"" Value=""#242730"" />
  <Setter Property=""Foreground"" Value=""#ECEEF3"" />
  <Setter Property=""BorderBrush"" Value=""#4B5260"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type Button}"">
        <Border x:Name=""ButtonBorder""
                Background=""{TemplateBinding Background}""
                BorderBrush=""{TemplateBinding BorderBrush}""
                BorderThickness=""{TemplateBinding BorderThickness}""
                Padding=""{TemplateBinding Padding}""
                CornerRadius=""5""
                SnapsToDevicePixels=""True"">
          <ContentPresenter HorizontalAlignment=""Center""
                            VerticalAlignment=""Center""
                            RecognizesAccessKey=""True"" />
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#323743"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#778191"" />
          </Trigger>
          <Trigger Property=""IsKeyboardFocused"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#263C50"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#56C4F5"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
          <Trigger Property=""IsPressed"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#1476AD"" />
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

        private static Style MakePrimaryButtonStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type Button}"">
  <Setter Property=""Background"" Value=""#168ED1"" />
  <Setter Property=""Foreground"" Value=""#FFFFFF"" />
  <Setter Property=""BorderBrush"" Value=""#168ED1"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type Button}"">
        <Border x:Name=""ButtonBorder""
                Background=""{TemplateBinding Background}""
                BorderBrush=""{TemplateBinding BorderBrush}""
                BorderThickness=""{TemplateBinding BorderThickness}""
                Padding=""{TemplateBinding Padding}""
                CornerRadius=""5""
                SnapsToDevicePixels=""True"">
          <ContentPresenter HorizontalAlignment=""Center""
                            VerticalAlignment=""Center""
                            RecognizesAccessKey=""True"" />
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#20A2E6"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#56C4F5"" />
          </Trigger>
          <Trigger Property=""IsKeyboardFocused"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""BorderBrush"" Value=""#B8E9FF"" />
            <Setter TargetName=""ButtonBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
          <Trigger Property=""IsPressed"" Value=""True"">
            <Setter TargetName=""ButtonBorder"" Property=""Background"" Value=""#1476AD"" />
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

        private static Style MakeToggleStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type CheckBox}"">
  <Setter Property=""Foreground"" Value=""#ECEEF3"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type CheckBox}"">
        <Border x:Name=""FocusBorder""
                Background=""Transparent""
                BorderBrush=""Transparent""
                BorderThickness=""2""
                CornerRadius=""6""
                Padding=""5,4""
                SnapsToDevicePixels=""True"">
          <StackPanel Orientation=""Horizontal"">
            <Border x:Name=""ToggleTrack""
                    Width=""54""
                    Height=""30""
                    Background=""#3A3F4A""
                    BorderBrush=""#687282""
                    BorderThickness=""1""
                    CornerRadius=""15""
                    VerticalAlignment=""Center"">
              <Ellipse x:Name=""ToggleKnob""
                       Width=""22""
                       Height=""22""
                       Fill=""#E8EBF0""
                       HorizontalAlignment=""Left""
                       VerticalAlignment=""Center""
                       Margin=""3,0"" />
            </Border>
            <ContentPresenter Margin=""10,0,0,0""
                              VerticalAlignment=""Center""
                              RecognizesAccessKey=""True"" />
          </StackPanel>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsChecked"" Value=""True"">
            <Setter TargetName=""ToggleTrack"" Property=""Background"" Value=""#168ED1"" />
            <Setter TargetName=""ToggleTrack"" Property=""BorderBrush"" Value=""#56C4F5"" />
            <Setter TargetName=""ToggleKnob"" Property=""HorizontalAlignment"" Value=""Right"" />
          </Trigger>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""FocusBorder"" Property=""Background"" Value=""#30343E"" />
          </Trigger>
          <Trigger Property=""IsKeyboardFocused"" Value=""True"">
            <Setter TargetName=""FocusBorder"" Property=""Background"" Value=""#263C50"" />
            <Setter TargetName=""FocusBorder"" Property=""BorderBrush"" Value=""#56C4F5"" />
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
