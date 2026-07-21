using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
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
        private static readonly Brush Accent = BrushFrom("#1683D8");

        private readonly PluginConfiguration _configuration;
        private readonly CheckBox _translatorEnabled;
        private readonly CheckBox _joyToKeyEnabled;
        private readonly ComboBox _profile;
        private readonly Button _refresh;
        private readonly Button _browseTranslator;
        private readonly Button _browseJoyToKey;
        private readonly Button _browseProfiles;
        private readonly Button _save;
        private readonly Button _cancel;
        private readonly TextBlock _profileStatus;
        private readonly TextBlock _readiness;
        private readonly List<Control> _focusOrder;

        public GameConfiguration Result { get; private set; }

        public GameSetupWindow(PluginConfiguration configuration, GameConfiguration game)
        {
            _configuration = configuration;
            Result = game;

            Title = "JRPG Translator Setup";
            Width = 760;
            Height = 590;
            MinWidth = 680;
            MinHeight = 540;
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
            options.Children.Add(_translatorEnabled);

            _joyToKeyEnabled = MakeCheckBox("Switch JoyToKey profile for this game", game.JoyToKeyEnabled);
            _joyToKeyEnabled.Margin = new Thickness(18, 10, 18, 14);
            _joyToKeyEnabled.Checked += (_, _) => UpdateJoyToKeyControls();
            _joyToKeyEnabled.Unchecked += (_, _) => UpdateJoyToKeyControls();
            options.Children.Add(_joyToKeyEnabled);

            Grid profileRow = new Grid { Margin = new Thickness(54, 0, 18, 10) };
            profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            options.Children.Add(profileRow);

            TextBlock profileLabel = new TextBlock
            {
                Text = "Profile:",
                Foreground = PrimaryForeground,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 14, 0)
            };
            profileRow.Children.Add(profileLabel);

            _profile = new ComboBox
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
            Grid.SetColumn(_profile, 1);
            profileRow.Children.Add(_profile);

            _refresh = MakeButton("Refresh", 116);
            _refresh.Margin = new Thickness(12, 0, 0, 0);
            _refresh.Click += (_, _) => RefreshProfiles(_profile.Text);
            Grid.SetColumn(_refresh, 2);
            profileRow.Children.Add(_refresh);

            _profileStatus = new TextBlock
            {
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(54, 0, 18, 12)
            };
            options.Children.Add(_profileStatus);

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

            _browseProfiles = MakeButton("Profiles Folder...", 162);
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

            _focusOrder = new List<Control>
            {
                _translatorEnabled,
                _joyToKeyEnabled,
                _profile,
                _refresh,
                _browseTranslator,
                _browseJoyToKey,
                _browseProfiles,
                _save,
                _cancel
            };

            PreviewKeyDown += HandlePreviewKeyDown;
            Loaded += (_, _) =>
            {
                // LaunchBox can keep its busy cursor active while a synchronous
                // plugin command is open. This dialog has no background work.
                Mouse.OverrideCursor = null;
                Cursor = Cursors.Arrow;
                ForceCursor = true;
                _translatorEnabled.Focus();
            };

            RefreshProfiles(game.JoyToKeyProfile);
            UpdateJoyToKeyControls();
        }

        private void RefreshProfiles(string preferredProfile)
        {
            IReadOnlyList<string> profiles = JoyToKeyProfileDiscovery.GetProfiles(
                _configuration.JoyToKeyProfilesDirectory);
            _profile.ItemsSource = profiles;

            string? selected = profiles.FirstOrDefault(
                value => string.Equals(value, preferredProfile, StringComparison.OrdinalIgnoreCase));
            _profile.SelectedItem = selected ?? profiles.FirstOrDefault();

            _profileStatus.Text = profiles.Count == 0
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
            _profile.IsEnabled = enabled;
            _refresh.IsEnabled = enabled;
        }

        private void BrowseTranslatorClicked(object sender, RoutedEventArgs e)
        {
            string current = PluginPaths.ResolveTranslatorExecutable(_configuration);
            string? selected = BrowseForExecutable(
                "Select JRPG Translator.exe",
                current);
            if (string.IsNullOrWhiteSpace(selected))
            {
                return;
            }

            _configuration.TranslatorExecutable = MakeLaunchBoxRelative(selected);
            UpdatePathStatus();
        }

        private void BrowseJoyToKeyClicked(object sender, RoutedEventArgs e)
        {
            string? selected = BrowseForExecutable(
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
            if (dialog.ShowDialog(this) != true || string.IsNullOrWhiteSpace(dialog.FolderName))
            {
                return;
            }

            _configuration.JoyToKeyProfilesDirectory = dialog.FolderName;
            RefreshProfiles(_profile.Text);
            UpdatePathStatus();
        }

        private static string? BrowseForExecutable(string title, string currentPath)
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
            return dialog.ShowDialog() == true ? dialog.FileName : null;
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
            if (_joyToKeyEnabled.IsChecked == true && _profile.SelectedItem == null)
            {
                MessageBox.Show(
                    this,
                    "Select a JoyToKey profile, or turn off JoyToKey profile switching for this game.",
                    "JRPG Translator Setup",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                _profile.Focus();
                return;
            }

            Result.TranslatorEnabled = _translatorEnabled.IsChecked == true;
            Result.JoyToKeyEnabled = _joyToKeyEnabled.IsChecked == true;
            Result.JoyToKeyProfile = _profile.SelectedItem as string ?? string.Empty;
            DialogResult = true;
        }

        private void HandlePreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (_profile.IsKeyboardFocusWithin && _profile.IsDropDownOpen)
            {
                if (e.Key == Key.Enter || e.Key == Key.Escape)
                {
                    _profile.IsDropDownOpen = false;
                    e.Handled = true;
                }
                return;
            }

            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                MoveFocus(e.Key == Key.Down ? 1 : -1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Left || e.Key == Key.Right)
            {
                if (_save.IsKeyboardFocused || _cancel.IsKeyboardFocused)
                {
                    (_save.IsKeyboardFocused ? _cancel : _save).Focus();
                    e.Handled = true;
                }
                else if (_browseTranslator.IsKeyboardFocused
                    || _browseJoyToKey.IsKeyboardFocused
                    || _browseProfiles.IsKeyboardFocused)
                {
                    Control[] locationControls =
                    {
                        _browseTranslator,
                        _browseJoyToKey,
                        _browseProfiles
                    };
                    int current = Array.FindIndex(
                        locationControls,
                        control => control.IsKeyboardFocused);
                    int direction = e.Key == Key.Right ? 1 : -1;
                    int next = (current + direction + locationControls.Length) % locationControls.Length;
                    locationControls[next].Focus();
                    e.Handled = true;
                }
                return;
            }

            if (e.Key != Key.Enter && e.Key != Key.Space)
            {
                return;
            }

            if (_translatorEnabled.IsKeyboardFocused)
            {
                _translatorEnabled.IsChecked = _translatorEnabled.IsChecked != true;
                e.Handled = true;
            }
            else if (_joyToKeyEnabled.IsKeyboardFocused)
            {
                _joyToKeyEnabled.IsChecked = _joyToKeyEnabled.IsChecked != true;
                e.Handled = true;
            }
            else if (_profile.IsKeyboardFocusWithin)
            {
                _profile.IsDropDownOpen = true;
                e.Handled = true;
            }
            else if (_refresh.IsKeyboardFocused)
            {
                _refresh.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
            else if (_browseTranslator.IsKeyboardFocused)
            {
                _browseTranslator.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
            else if (_browseJoyToKey.IsKeyboardFocused)
            {
                _browseJoyToKey.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
            else if (_browseProfiles.IsKeyboardFocused)
            {
                _browseProfiles.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
            else if (_save.IsKeyboardFocused)
            {
                _save.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
            else if (_cancel.IsKeyboardFocused)
            {
                _cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                e.Handled = true;
            }
        }

        private void MoveFocus(int direction)
        {
            int current = _focusOrder.FindIndex(control => control.IsKeyboardFocusWithin);
            if (current < 0)
            {
                current = 0;
            }

            for (int offset = 1; offset <= _focusOrder.Count; offset++)
            {
                int next = (current + (offset * direction) + _focusOrder.Count) % _focusOrder.Count;
                Control candidate = _focusOrder[next];
                if (candidate.IsEnabled && candidate.Visibility == Visibility.Visible)
                {
                    candidate.Focus();
                    return;
                }
            }
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
                FocusVisualStyle = MakeFocusStyle()
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
                FocusVisualStyle = MakeFocusStyle()
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

        private static Style MakeFocusStyle()
        {
            Style style = new Style(typeof(Control));
            ControlTemplate template = new ControlTemplate(typeof(Control));
            FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.BorderBrushProperty, Accent);
            border.SetValue(Border.BorderThicknessProperty, new Thickness(2));
            border.SetValue(Border.MarginProperty, new Thickness(-3));
            template.VisualTree = border;
            style.Setters.Add(new Setter(Control.TemplateProperty, template));
            return style;
        }

        private static SolidColorBrush BrushFrom(string value)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value));
        }
    }
}
