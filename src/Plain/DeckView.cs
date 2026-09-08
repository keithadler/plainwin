using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Plain.Core;

namespace Plain;

/// <summary>
/// A deck as a rail of slides and one slide's text, editable. Plain does not draw shapes, pictures or diagrams; it
/// shows what text is on the slide and says what else the slide is holding.
/// </summary>
public sealed class DeckView : Grid
{
    private readonly Deck _deck;
    private readonly StackPanel _rail = new() { Margin = new Thickness(8, 10, 8, 10) };
    private readonly StackPanel _canvas = new() { Margin = new Thickness(30, 26, 30, 30), MaxWidth = 720 };
    private Slide _current;

    public event Action<Action>? Edited;
    public Slide Current => _current;
    public bool FlattenedSomething { get; private set; }

    public DeckView(Deck deck)
    {
        _deck = deck;
        _current = deck.Slides.Count > 0 ? deck.Slides[0] : throw new OpcPackage.PackageException("This deck has no slides.");

        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var railScroll = new ScrollViewer
        {
            Content = _rail,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Background = App.B("Surface2"),
        };
        var railBorder = new Border { Child = railScroll, BorderBrush = App.B("Line"), BorderThickness = new Thickness(0, 0, 1, 0) };
        SetColumn(railBorder, 0); Children.Add(railBorder);

        var canvasScroll = new ScrollViewer
        {
            Content = _canvas,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Background = App.B("Surface"),
        };
        SetColumn(canvasScroll, 1); Children.Add(canvasScroll);

        BuildRail();
        ShowSlide(_current);
    }

    private void BuildRail()
    {
        _rail.Children.Clear();
        foreach (var slide in _deck.Slides)
        {
            var number = new TextBlock
            {
                Text = slide.Number.ToString(),
                FontFamily = new FontFamily("Cascadia Mono, Consolas"),
                FontSize = 10,
                Foreground = App.B("Ink3"),
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 0, 6, 0),
            };
            var title = new TextBlock
            {
                Text = slide.Title(),
                FontSize = 10.5,
                TextTrimming = TextTrimming.CharacterEllipsis,
                TextWrapping = TextWrapping.Wrap,
                MaxHeight = 40,
                Foreground = App.B("Ink2"),
            };
            var content = new DockPanel();
            DockPanel.SetDock(number, Dock.Left);
            content.Children.Add(number);
            content.Children.Add(title);

            var button = new Button
            {
                Content = content,
                Margin = new Thickness(0, 0, 0, 7),
                Padding = new Thickness(7, 6, 7, 6),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Background = App.B("Surface"),
                BorderBrush = slide == _current ? App.B("Accent") : App.B("LineStrong"),
                BorderThickness = new Thickness(slide == _current ? 2 : 1),
                Cursor = System.Windows.Input.Cursors.Hand,
                Tag = slide,
            };
            button.Click += (s, _) => { _current = (Slide)((Button)s).Tag; BuildRail(); ShowSlide(_current); };
            _rail.Children.Add(button);
        }
    }

    private void ShowSlide(Slide slide)
    {
        _canvas.Children.Clear();
        foreach (var frame in slide.Texts())
        {
            var label = new TextBlock
            {
                Text = Describe(frame.Placeholder),
                FontSize = 10.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = App.B("Ink3"),
                Margin = new Thickness(0, 14, 0, 4),
            };
            _canvas.Children.Add(label);

            for (int i = 0; i < frame.Lines.Count; i++)
            {
                int lineIndex = i;
                int shapeIndex = frame.ShapeIndex;
                bool isTitle = frame.Placeholder is "title" or "ctrTitle";
                var box = new TextBox
                {
                    Text = frame.Lines[i],
                    FontSize = isTitle ? 20 : 13.5,
                    FontWeight = isTitle ? FontWeights.SemiBold : FontWeights.Normal,
                    BorderThickness = new Thickness(0),
                    Background = Brushes.Transparent,
                    Foreground = App.B("Ink"),
                    Padding = new Thickness(2, 2, 2, 2),
                    TextWrapping = TextWrapping.Wrap,
                    AcceptsReturn = false,
                };
                string was = frame.Lines[i];
                box.LostFocus += (s, _) =>
                {
                    var edited = (TextBox)s;
                    if (edited.Text == was) return;
                    if (!slide.SetLine(shapeIndex, lineIndex, edited.Text)) FlattenedSomething = true;
                    BuildRail();
                    Edited?.Invoke(() =>
                    {
                        slide.SetLine(shapeIndex, lineIndex, was);
                        BuildRail();
                        ShowSlide(slide);
                    });
                };
                _canvas.Children.Add(box);
            }
        }

        var note = new TextBlock
        {
            Text = "Pictures, diagrams and shapes on this slide are kept in the file and left exactly where they are.",
            FontSize = 11,
            TextWrapping = TextWrapping.Wrap,
            Foreground = App.B("Preserved"),
            Margin = new Thickness(0, 24, 0, 0),
        };
        _canvas.Children.Add(note);
    }

    /// <summary>
    /// What to call a text frame. Some programs write a slide with no placeholder type at all, and calling every one
    /// of those "body" would be a guess; they are simply text.
    /// </summary>
    private static string Describe(string placeholder) => placeholder switch
    {
        "title" or "ctrTitle" => "TITLE",
        "subTitle" => "SUBTITLE",
        "body" => "BODY",
        "ftr" => "FOOTER",
        "sldNum" => "SLIDE NUMBER",
        "" => "TEXT",
        _ => placeholder.ToUpperInvariant(),
    };
}
