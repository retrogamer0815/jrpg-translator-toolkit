using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using JrpgTranslator.LaunchBox;
using Unbroken.LaunchBox.Plugins.Data;

internal static class CoverArtworkTests
{
    private static int assertions;

    internal static void Run(PluginConfiguration configuration, GameConfiguration original,
        string fixtureDirectory, string? renderDirectory)
    {
        assertions = 0;
        string portrait = Path.Combine(fixtureDirectory, "cover #1 日本語 %.png");
        string square = Path.Combine(fixtureDirectory, "square.jpg");
        string landscape = Path.Combine(fixtureDirectory, "landscape.png");
        string corrupt = Path.Combine(fixtureDirectory, "corrupt.png");
        string locked = Path.Combine(fixtureDirectory, "locked.png");
        WriteCover(portrait, 640, 960);
        WriteCover(square, 900, 900);
        WriteCover(landscape, 1200, 600);
        File.WriteAllText(corrupt, "Not an image");
        File.Copy(portrait, locked);
        using FileStream heldLock = File.Open(locked, FileMode.Open, FileAccess.Read, FileShare.None);

        IGame selectedGame = DispatchProxy.Create<IGame, CoverGameProxy>();
        CoverGameProxy proxy = (CoverGameProxy)selectedGame;
        proxy.CoverPath = portrait;
        MethodInfo getPath = typeof(GameSetupMenuItem).GetMethod("GetCoverArtPath",
            BindingFlags.Static | BindingFlags.NonPublic)!;
        Check((string?)getPath.Invoke(null, new object[] { selectedGame }) == portrait,
            "Setup must read the selected game's front-cover property.");
        proxy.ThrowOnRead = true;
        Check(getPath.Invoke(null, new object[] { selectedGame }) == null,
            "Artwork lookup failure must not prevent setup from opening.");

        var cases = new[]
        {
            (Name: "portrait", Path: (string?)portrait, HasArt: true, LongTitle: false),
            (Name: "square", Path: (string?)square, HasArt: true, LongTitle: false),
            (Name: "landscape", Path: (string?)landscape, HasArt: true, LongTitle: false),
            (Name: "long-title", Path: (string?)portrait, HasArt: true, LongTitle: true),
            (Name: "no-cover", Path: (string?)null, HasArt: false, LongTitle: false),
            (Name: "blank", Path: (string?)"   ", HasArt: false, LongTitle: false),
            (Name: "missing", Path: (string?)Path.Combine(fixtureDirectory, "missing.png"), HasArt: false, LongTitle: false),
            (Name: "corrupt", Path: (string?)corrupt, HasArt: false, LongTitle: false),
            (Name: "locked", Path: (string?)locked, HasArt: false, LongTitle: false)
        };
        if (renderDirectory != null) Directory.CreateDirectory(renderDirectory);
        foreach (var sample in cases)
        {
            GameConfiguration game = original.Clone();
            game.GameTitle = sample.LongTitle
                ? "A Very Long Game Title: 日本語 — An Extended Edition with an Extra Subtitle that Cannot Fit in a Small Window"
                : "The Legend of Heroes III: Shiroki Majo";
            GameSetupWindow window = new GameSetupWindow(configuration, game, sample.Path);
            Grid root = (Grid)window.Content;
            Grid header = root.Children.OfType<Grid>().Single();
            StackPanel text = header.Children.OfType<StackPanel>().Single();
            TextBlock title = text.Children.OfType<TextBlock>().First();
            TextBlock description = text.Children.OfType<TextBlock>().Last();
            Border? art = header.Children.OfType<Border>().SingleOrDefault();
            ScrollViewer scroller = root.Children.OfType<ScrollViewer>().Single();
            StackPanel footer = root.Children.OfType<StackPanel>().Single();
            Check(window.Icon == null, "The restored title-bar branding must remain unchanged.");
            Check((art != null) == sample.HasArt, sample.Name + ": incorrect artwork fallback.");
            Check(title.TextTrimming == TextTrimming.CharacterEllipsis
                && title.ToolTip as string == game.GameTitle && description.TextWrapping == TextWrapping.Wrap,
                "Header text must trim/wrap safely and expose the full title on hover.");
            if (art != null)
            {
                Image image = (Image)art.Child;
                BitmapSource source = (BitmapSource)image.Source;
                Check(source.IsFrozen && source.PixelWidth <= 360 && source.PixelHeight <= 312,
                    "Cover decoding must be bounded and independent of the file stream.");
                Check(image.Stretch == Stretch.Uniform && image.Width <= 120.1 && image.Height <= 104.1
                    && Math.Abs(image.Width / image.Height - (double)source.PixelWidth / source.PixelHeight) < 0.001,
                    "Cover art must fit without stretching or cropping.");
                Check(!art.Focusable && !art.IsHitTestVisible,
                    "Decorative artwork must not intercept input.");
                using FileStream exclusive = File.Open(sample.Path!, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                Check(exclusive.CanWrite, "Showing a cover must not lock the LaunchBox artwork file.");
            }

            // A finite offscreen host models the viewport without opening native
            // windows. Keep inherited typography identical to the real window.
            window.Content = null;
            ContentControl surface = new ContentControl
            {
                Content = root, Background = window.Background, Foreground = window.Foreground,
                FontFamily = window.FontFamily, FontSize = window.FontSize,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment = VerticalAlignment.Stretch
            };
            try
            {
                foreach (var size in new[] { (Width: 920, Height: 800), (Width: 760, Height: 560), (Width: 1280, Height: 800) })
                {
                    surface.Width = size.Width;
                    surface.Height = size.Height;
                    surface.Measure(new Size(size.Width, size.Height));
                    surface.Arrange(new Rect(0, 0, size.Width, size.Height));
                    surface.UpdateLayout();
                    Rect headerBounds = Bounds(header, root);
                    Rect textBounds = Bounds(text, root);
                    Rect scrollBounds = Bounds(scroller, root);
                    Rect footerBounds = Bounds(footer, root);
                    Check(root.ActualWidth <= size.Width - 60 + 0.1, "Content exceeds viewport width.");
                    Check(Math.Abs(textBounds.Left - headerBounds.Left) < 0.1, "Text must stay left-aligned.");
                    Check(Bounds(description, root).Top >= Bounds(title, root).Bottom + 7,
                        "Title and description overlap.");
                    Check(scrollBounds.Top >= headerBounds.Bottom + 19 && scrollBounds.Height > 140,
                        "The header must leave usable scrolling content below it.");
                    Check(footerBounds.Top >= scrollBounds.Bottom - 0.1 && footerBounds.Bottom <= root.ActualHeight + 0.1,
                        "The Save/Cancel footer must remain visible outside scrolling content.");
                    if (art != null)
                    {
                        Rect artBounds = Bounds(art, root);
                        Check(artBounds.Left >= textBounds.Right + 23
                            && Math.Abs(artBounds.Right - headerBounds.Right) < 0.1,
                            "The cover must remain at the upper-right, clear of the text.");
                        Check(Math.Abs(artBounds.Top - headerBounds.Top) < 0.1,
                            "Artwork must be aligned to the top of the header.");
                    }
                    else
                    {
                        Check(Math.Abs(textBounds.Width - headerBounds.Width) < 0.1,
                            "Absent/unreadable covers must leave no reserved column or margin.");
                    }
                    if (renderDirectory != null && (sample.HasArt || sample.Name == "no-cover"))
                    {
                        foreach (double scale in new[] { 1.0, 1.5, 2.0 })
                            Render(surface, window.Background, scale,
                                Path.Combine(renderDirectory, $"{sample.Name}-{size.Width}-{scale * 100:0}.png"));
                    }
                }
            }
            finally
            {
                surface.Content = null;
                window.Content = root;
                window.Close();
            }
        }
        Console.WriteLine($"Cover artwork passed: {assertions} assertions; 9 artwork/title cases at 3 viewport sizes.");
    }

    private static Rect Bounds(FrameworkElement element, Visual relativeTo) =>
        element.TransformToAncestor(relativeTo).TransformBounds(new Rect(element.RenderSize));

    private static void WriteCover(string path, int width, int height)
    {
        DrawingVisual drawing = new DrawingVisual();
        using (DrawingContext context = drawing.RenderOpen())
        {
            context.DrawRectangle(new SolidColorBrush(Color.FromRgb(30, 72, 108)), null, new Rect(0, 0, width, height));
            context.DrawRectangle(new SolidColorBrush(Color.FromRgb(229, 190, 99)), null,
                new Rect(width * 0.08, height * 0.3, width * 0.84, height * 0.6));
            FormattedText label = new FormattedText("TEST COVER", CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight, new Typeface("Segoe UI"), width * 0.10, Brushes.White, 1);
            context.DrawText(label, new Point(width * 0.08, height * 0.08));
        }
        RenderTargetBitmap image = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        image.Render(drawing);
        BitmapEncoder encoder = path.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
            ? new JpegBitmapEncoder() : new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        using FileStream file = File.Create(path);
        encoder.Save(file);
    }

    private static void Render(FrameworkElement element, Brush background, double scale, string path)
    {
        DrawingVisual drawing = new DrawingVisual();
        Rect bounds = new Rect(0, 0, element.ActualWidth, element.ActualHeight);
        using (DrawingContext context = drawing.RenderOpen())
        {
            context.DrawRectangle(background, null, bounds);
            context.DrawRectangle(new VisualBrush(element)
            {
                AutoLayoutContent = false, ViewboxUnits = BrushMappingMode.Absolute, Viewbox = bounds
            }, null, bounds);
        }
        RenderTargetBitmap image = new RenderTargetBitmap((int)Math.Ceiling(bounds.Width * scale),
            (int)Math.Ceiling(bounds.Height * scale), 96 * scale, 96 * scale, PixelFormats.Pbgra32);
        image.Render(drawing);
        PngBitmapEncoder encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(image));
        using FileStream file = File.Create(path);
        encoder.Save(file);
    }

    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new InvalidOperationException(message);
    }
}

public class CoverGameProxy : DispatchProxy
{
    public string? CoverPath { get; set; }
    public bool ThrowOnRead { get; set; }
    protected override object? Invoke(MethodInfo? method, object?[]? args)
    {
        if (method?.Name != "get_FrontImagePath") throw new InvalidOperationException("Unexpected game lookup");
        if (ThrowOnRead) throw new IOException("Synthetic artwork lookup failure");
        return CoverPath;
    }
}
