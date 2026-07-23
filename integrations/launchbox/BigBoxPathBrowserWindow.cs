using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace JrpgTranslator.LaunchBox
{
    internal enum BigBoxPathBrowserMode
    {
        Executable,
        Folder
    }

    internal sealed class BigBoxPathBrowserWindow : Window
    {
        private static readonly Brush WindowBackground = BrushFrom("#202124");
        private static readonly Brush ControlBackground = BrushFrom("#333438");
        private static readonly Brush PrimaryForeground = BrushFrom("#F3F4F6");
        private static readonly Brush MutedForeground = BrushFrom("#AEB3BC");
        private static readonly Brush ControlBorderBrush = BrushFrom("#686B72");

        private readonly BigBoxPathBrowserMode _mode;
        private readonly TextBlock _currentPath;
        private readonly ListBox _entries;
        private readonly TextBlock _status;
        private readonly Button _select;
        private readonly Button _parent;
        private readonly Button _cancel;
        private readonly Control[] _buttonRow;
        private readonly DispatcherTimer _controllerTimer;
        private readonly Dictionary<ControllerNavigationCommand, long> _keyboardMirrorGraceUntil = new();
        private readonly Dictionary<ControllerNavigationCommand, long> _lastNativeNavigationAt = new();
        private ControllerNavigationState _controllerPreviousState;
        private ControllerNavigationCommand? _heldControllerDirection;
        private long _nextControllerRepeatAt;
        private bool _controllerBaselineReady;
        private string _currentDirectory = string.Empty;

        private BigBoxPathBrowserWindow(
            string title,
            string initialPath,
            BigBoxPathBrowserMode mode)
        {
            _mode = mode;

            Title = title;
            Width = 820;
            Height = 680;
            MinWidth = 620;
            MinHeight = 500;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
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
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            TextBlock heading = new TextBlock
            {
                Text = mode == BigBoxPathBrowserMode.Executable
                    ? "Choose an application"
                    : "Choose a folder",
                FontSize = 25,
                FontWeight = FontWeights.SemiBold,
                Foreground = PrimaryForeground,
                Margin = new Thickness(0, 0, 0, 8)
            };
            root.Children.Add(heading);

            TextBlock instructions = new TextBlock
            {
                Text = mode == BigBoxPathBrowserMode.Executable
                    ? "A opens folders or selects an .exe file. B returns to the parent folder."
                    : "A opens folders. Use Select this folder to choose the displayed folder. B returns to the parent folder.",
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 18)
            };
            Grid.SetRow(instructions, 1);
            root.Children.Add(instructions);

            Border pathBorder = new Border
            {
                Background = ControlBackground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 12)
            };
            _currentPath = new TextBlock
            {
                Foreground = PrimaryForeground,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            pathBorder.Child = _currentPath;
            Grid.SetRow(pathBorder, 2);
            root.Children.Add(pathBorder);

            _entries = new ListBox
            {
                Background = ControlBackground,
                Foreground = PrimaryForeground,
                BorderBrush = ControlBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(2),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Style = MakeListBoxStyle(),
                ItemContainerStyle = MakeListBoxItemStyle()
            };
            ScrollViewer.SetVerticalScrollBarVisibility(_entries, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(_entries, ScrollBarVisibility.Disabled);
            _entries.MouseDoubleClick += (_, _) => ActivateSelectedEntry();
            _entries.SelectionChanged += (_, _) => UpdateSelectionState();
            Grid.SetRow(_entries, 3);
            root.Children.Add(_entries);

            _status = new TextBlock
            {
                Foreground = MutedForeground,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 12, 0, 14)
            };
            Grid.SetRow(_status, 4);
            root.Children.Add(_status);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            Grid.SetRow(buttons, 5);
            root.Children.Add(buttons);

            _select = MakeButton(
                mode == BigBoxPathBrowserMode.Executable ? "Select application" : "Select this folder",
                190);
            _select.Click += (_, _) => SelectCurrentChoice();
            buttons.Children.Add(_select);

            _parent = MakeButton("Parent folder", 160);
            _parent.Margin = new Thickness(12, 0, 0, 0);
            _parent.Click += (_, _) => GoToParentOrDrives();
            buttons.Children.Add(_parent);

            _cancel = MakeButton("Cancel", 130);
            _cancel.Margin = new Thickness(12, 0, 0, 0);
            _cancel.IsCancel = true;
            _cancel.Click += (_, _) => DialogResult = false;
            buttons.Children.Add(_cancel);

            _buttonRow = new Control[] { _select, _parent, _cancel };
            _controllerTimer = new DispatcherTimer(DispatcherPriority.Input)
            {
                Interval = TimeSpan.FromMilliseconds(20)
            };
            _controllerTimer.Tick += HandleControllerTick;

            PreviewKeyDown += HandlePreviewKeyDown;
            PreviewKeyUp += HandlePreviewKeyUp;
            Loaded += (_, _) =>
            {
                Mouse.OverrideCursor = null;
                Cursor = Cursors.Arrow;
                ForceCursor = true;
                _controllerTimer.Start();
                _entries.Focus();
            };
            Activated += (_, _) => ResetControllerNavigation();
            Deactivated += (_, _) => ResetControllerNavigation();
            Closed += (_, _) => _controllerTimer.Stop();

            NavigateTo(ResolveInitialDirectory(initialPath, mode));
        }

        internal string? SelectedPath { get; private set; }

        internal static string? BrowseExecutable(Window owner, string title, string currentPath)
        {
            BigBoxPathBrowserWindow browser = new BigBoxPathBrowserWindow(
                title,
                currentPath,
                BigBoxPathBrowserMode.Executable)
            {
                Owner = owner
            };
            return browser.ShowDialog() == true ? browser.SelectedPath : null;
        }

        internal static string? BrowseFolder(Window owner, string title, string currentPath)
        {
            BigBoxPathBrowserWindow browser = new BigBoxPathBrowserWindow(
                title,
                currentPath,
                BigBoxPathBrowserMode.Folder)
            {
                Owner = owner
            };
            return browser.ShowDialog() == true ? browser.SelectedPath : null;
        }

        private static string ResolveInitialDirectory(string initialPath, BigBoxPathBrowserMode mode)
        {
            if (!string.IsNullOrWhiteSpace(initialPath))
            {
                if (Directory.Exists(initialPath))
                {
                    return Path.GetFullPath(initialPath);
                }

                string? directory = mode == BigBoxPathBrowserMode.Executable || Path.HasExtension(initialPath)
                    ? Path.GetDirectoryName(initialPath)
                    : initialPath;
                if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                {
                    return Path.GetFullPath(directory);
                }
            }

            if (!string.IsNullOrWhiteSpace(PluginPaths.LaunchBoxRoot)
                && Directory.Exists(PluginPaths.LaunchBoxRoot))
            {
                return PluginPaths.LaunchBoxRoot;
            }

            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Directory.Exists(documents) ? documents : string.Empty;
        }

        private void NavigateTo(string directory)
        {
            _currentDirectory = directory;
            IReadOnlyList<BrowserEntry> entries;
            string message;

            if (string.IsNullOrWhiteSpace(directory))
            {
                entries = LoadDrives();
                message = entries.Count == 0
                    ? "No ready drives were found."
                    : entries.Count + (entries.Count == 1 ? " drive" : " drives");
                _currentPath.Text = "Computer";
            }
            else
            {
                try
                {
                    string[] directories = Directory.GetDirectories(directory);
                    IEnumerable<BrowserEntry> folders = directories
                        .OrderBy(path => Path.GetFileName(path), StringComparer.CurrentCultureIgnoreCase)
                        .Select(path => new BrowserEntry(path, "Folder", Path.GetFileName(path)));

                    IEnumerable<BrowserEntry> files = Enumerable.Empty<BrowserEntry>();
                    if (_mode == BigBoxPathBrowserMode.Executable)
                    {
                        files = Directory.GetFiles(directory, "*.exe")
                            .OrderBy(path => Path.GetFileName(path), StringComparer.CurrentCultureIgnoreCase)
                            .Select(path => new BrowserEntry(path, "Application", Path.GetFileName(path)));
                    }

                    entries = folders.Concat(files).ToArray();
                    message = entries.Count + (entries.Count == 1 ? " item" : " items");
                    _currentPath.Text = directory;
                }
                catch (Exception exception) when (
                    exception is UnauthorizedAccessException
                    || exception is IOException
                    || exception is System.Security.SecurityException)
                {
                    entries = Array.Empty<BrowserEntry>();
                    message = "This folder cannot be opened. Use B or Parent folder to go back.";
                    _currentPath.Text = directory;
                }
            }

            _entries.ItemsSource = entries;
            _entries.SelectedIndex = entries.Count > 0 ? 0 : -1;
            _status.Text = message;
            UpdateSelectionState();
            _entries.ScrollIntoView(_entries.SelectedItem);
        }

        private static IReadOnlyList<BrowserEntry> LoadDrives()
        {
            List<BrowserEntry> drives = new List<BrowserEntry>();
            foreach (DriveInfo drive in DriveInfo.GetDrives().OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (!drive.IsReady)
                    {
                        continue;
                    }

                    string label = string.IsNullOrWhiteSpace(drive.VolumeLabel)
                        ? drive.Name
                        : drive.Name + "  " + drive.VolumeLabel;
                    drives.Add(new BrowserEntry(drive.RootDirectory.FullName, "Drive", label));
                }
                catch (IOException)
                {
                    // Removable drives can disappear while the browser is open.
                }
                catch (UnauthorizedAccessException)
                {
                    // Skip drives Windows does not allow this process to inspect.
                }
            }
            return drives;
        }

        private void UpdateSelectionState()
        {
            BrowserEntry? entry = _entries.SelectedItem as BrowserEntry;
            _select.IsEnabled = _mode == BigBoxPathBrowserMode.Folder
                ? !string.IsNullOrWhiteSpace(_currentDirectory)
                : entry?.Kind == "Application";
            _parent.IsEnabled = !string.IsNullOrWhiteSpace(_currentDirectory);

            if (entry != null)
            {
                _status.Text = entry.FullPath;
            }
        }

        private void ActivateSelectedEntry()
        {
            if (_entries.SelectedItem is not BrowserEntry entry)
            {
                return;
            }

            if (entry.Kind == "Application")
            {
                Accept(entry.FullPath);
            }
            else
            {
                NavigateTo(entry.FullPath);
                _entries.Focus();
            }
        }

        private void SelectCurrentChoice()
        {
            if (_mode == BigBoxPathBrowserMode.Folder)
            {
                if (!string.IsNullOrWhiteSpace(_currentDirectory))
                {
                    Accept(_currentDirectory);
                }
                return;
            }

            if (_entries.SelectedItem is BrowserEntry entry && entry.Kind == "Application")
            {
                Accept(entry.FullPath);
            }
        }

        private void Accept(string path)
        {
            SelectedPath = path;
            DialogResult = true;
        }

        private void GoToParentOrDrives()
        {
            if (string.IsNullOrWhiteSpace(_currentDirectory))
            {
                DialogResult = false;
                return;
            }

            DirectoryInfo? parent = null;
            try
            {
                parent = Directory.GetParent(_currentDirectory);
            }
            catch (Exception exception) when (
                exception is ArgumentException
                || exception is IOException
                || exception is UnauthorizedAccessException)
            {
                // A malformed or disconnected path falls back to the drive list.
            }

            NavigateTo(parent?.FullName ?? string.Empty);
            _entries.Focus();
        }

        private void HandlePreviewKeyDown(object sender, KeyEventArgs e)
        {
            ControllerNavigationCommand? mirrored = ControllerCommandForKey(e.Key);
            if (mirrored.HasValue && IsMirroredControllerKey(mirrored.Value))
            {
                e.Handled = true;
                return;
            }

            switch (e.Key)
            {
                case Key.Up:
                    DispatchNavigation(ControllerNavigationCommand.Up, false);
                    e.Handled = true;
                    break;
                case Key.Down:
                    DispatchNavigation(ControllerNavigationCommand.Down, false);
                    e.Handled = true;
                    break;
                case Key.Left:
                    DispatchNavigation(ControllerNavigationCommand.Left, false);
                    e.Handled = true;
                    break;
                case Key.Right:
                    DispatchNavigation(ControllerNavigationCommand.Right, false);
                    e.Handled = true;
                    break;
                case Key.Back:
                    GoToParentOrDrives();
                    e.Handled = true;
                    break;
                case Key.Enter:
                case Key.Space:
                    DispatchNavigation(ControllerNavigationCommand.Activate, false);
                    e.Handled = true;
                    break;
                case Key.Escape:
                    DialogResult = false;
                    e.Handled = true;
                    break;
            }
        }

        private void HandlePreviewKeyUp(object sender, KeyEventArgs e)
        {
            ControllerNavigationCommand? mirrored = ControllerCommandForKey(e.Key);
            if (mirrored.HasValue && IsMirroredControllerKey(mirrored.Value))
            {
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
                DispatchNavigation(newDirection!.Value, true);
                _heldControllerDirection = newDirection.Value;
                _nextControllerRepeatAt = now + 340;
            }
            else if (_heldControllerDirection.HasValue
                && state.IsPressed(_heldControllerDirection.Value))
            {
                if (now >= _nextControllerRepeatAt)
                {
                    DispatchNavigation(_heldControllerDirection.Value, true);
                    _nextControllerRepeatAt = now + 90;
                }
            }
            else
            {
                _heldControllerDirection = null;
                _nextControllerRepeatAt = 0;
            }

            if (state.Activate && !_controllerPreviousState.Activate)
            {
                DispatchNavigation(ControllerNavigationCommand.Activate, true);
            }
            if (state.Cancel && !_controllerPreviousState.Cancel)
            {
                DispatchNavigation(ControllerNavigationCommand.Cancel, true);
            }

            _controllerPreviousState = state;
        }

        private void DispatchNavigation(ControllerNavigationCommand command, bool native)
        {
            if (native)
            {
                _lastNativeNavigationAt[command] = Environment.TickCount64;
            }

            switch (command)
            {
                case ControllerNavigationCommand.Up:
                    MoveVertical(-1);
                    break;
                case ControllerNavigationCommand.Down:
                    MoveVertical(1);
                    break;
                case ControllerNavigationCommand.Left:
                    if (_entries.IsKeyboardFocusWithin)
                    {
                        GoToParentOrDrives();
                    }
                    else
                    {
                        MoveButtonFocus(-1);
                    }
                    break;
                case ControllerNavigationCommand.Right:
                    if (_entries.IsKeyboardFocusWithin)
                    {
                        ActivateSelectedEntry();
                    }
                    else
                    {
                        MoveButtonFocus(1);
                    }
                    break;
                case ControllerNavigationCommand.Activate:
                    ActivateFocusedControl();
                    break;
                case ControllerNavigationCommand.Cancel:
                    GoToParentOrDrives();
                    break;
            }
        }

        private void MoveVertical(int direction)
        {
            if (_entries.IsKeyboardFocusWithin)
            {
                if (_entries.Items.Count == 0)
                {
                    FocusFirstAvailableButton();
                    return;
                }

                int current = _entries.SelectedIndex < 0 ? 0 : _entries.SelectedIndex;
                int next = current + direction;
                if (next >= 0 && next < _entries.Items.Count)
                {
                    _entries.SelectedIndex = next;
                    _entries.ScrollIntoView(_entries.SelectedItem);
                }
                else if (direction > 0)
                {
                    FocusFirstAvailableButton();
                }
                return;
            }

            if (direction < 0)
            {
                _entries.Focus();
                if (_entries.SelectedIndex < 0 && _entries.Items.Count > 0)
                {
                    _entries.SelectedIndex = 0;
                }
            }
        }

        private void MoveButtonFocus(int direction)
        {
            Control[] available = _buttonRow.Where(control => control.IsEnabled).ToArray();
            int current = Array.FindIndex(available, control => control.IsKeyboardFocusWithin);
            int next = current + direction;
            if (current >= 0 && next >= 0 && next < available.Length)
            {
                available[next].Focus();
            }
        }

        private void FocusFirstAvailableButton()
        {
            _buttonRow.FirstOrDefault(control => control.IsEnabled)?.Focus();
        }

        private void ActivateFocusedControl()
        {
            if (_entries.IsKeyboardFocusWithin)
            {
                ActivateSelectedEntry();
                return;
            }

            Button? focused = _buttonRow
                .OfType<Button>()
                .FirstOrDefault(button => button.IsKeyboardFocused && button.IsEnabled);
            focused?.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }

        private void ResetControllerNavigation()
        {
            _controllerPreviousState = default;
            _heldControllerDirection = null;
            _nextControllerRepeatAt = 0;
            _controllerBaselineReady = false;
        }

        private static Button MakeButton(string text, double width)
        {
            return new Button
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
          <ContentPresenter HorizontalAlignment=""Center"" VerticalAlignment=""Center"" />
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

        private static Style MakeListBoxStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type ListBox}"">
  <Setter Property=""Background"" Value=""#333438"" />
  <Setter Property=""Foreground"" Value=""#F3F4F6"" />
  <Setter Property=""BorderBrush"" Value=""#686B72"" />
  <Setter Property=""BorderThickness"" Value=""1"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type ListBox}"">
        <Border x:Name=""ListBorder""
                Background=""{TemplateBinding Background}""
                BorderBrush=""{TemplateBinding BorderBrush}""
                BorderThickness=""{TemplateBinding BorderThickness}""
                Padding=""{TemplateBinding Padding}"">
          <ScrollViewer Focusable=""False"" Padding=""0"">
            <ItemsPresenter />
          </ScrollViewer>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsKeyboardFocusWithin"" Value=""True"">
            <Setter TargetName=""ListBorder"" Property=""BorderBrush"" Value=""#4FB3FF"" />
            <Setter TargetName=""ListBorder"" Property=""BorderThickness"" Value=""2"" />
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>";
            return (Style)XamlReader.Parse(xaml);
        }

        private static Style MakeListBoxItemStyle()
        {
            const string xaml = @"
<Style xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
       xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml""
       TargetType=""{x:Type ListBoxItem}"">
  <Setter Property=""Background"" Value=""Transparent"" />
  <Setter Property=""Foreground"" Value=""#F3F4F6"" />
  <Setter Property=""HorizontalContentAlignment"" Value=""Stretch"" />
  <Setter Property=""Padding"" Value=""10,8"" />
  <Setter Property=""Template"">
    <Setter.Value>
      <ControlTemplate TargetType=""{x:Type ListBoxItem}"">
        <Border x:Name=""ItemBorder""
                Background=""{TemplateBinding Background}""
                Padding=""{TemplateBinding Padding}""
                SnapsToDevicePixels=""True"">
          <TextBlock Text=""{Binding DisplayName}""
                     Foreground=""{TemplateBinding Foreground}""
                     TextTrimming=""CharacterEllipsis"" />
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property=""IsMouseOver"" Value=""True"">
            <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#3D3F44"" />
          </Trigger>
          <Trigger Property=""IsSelected"" Value=""True"">
            <Setter TargetName=""ItemBorder"" Property=""Background"" Value=""#1683D8"" />
            <Setter Property=""Foreground"" Value=""#FFFFFF"" />
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

        private sealed class BrowserEntry
        {
            internal BrowserEntry(string fullPath, string kind, string name)
            {
                FullPath = fullPath;
                Kind = kind;
                DisplayName = "[" + kind + "]  " + name;
            }

            internal string FullPath { get; }
            internal string Kind { get; }
            public string DisplayName { get; }
        }
    }
}
