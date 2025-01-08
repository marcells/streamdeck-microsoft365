using System.Drawing;
using System.Drawing.Imaging;
using Svg;

namespace Elgato.Plugins.Microsoft365;

class IconCreator
{
    public IconCreator(string? badgeImageFilePath)
    {
        BadgeImageFilePath = badgeImageFilePath;
    }

    public string? BadgeImageFilePath { get; init; }

    public string CreateNotificationSvg(int count, Color foregroundColor, Color backgroundColor, string? header = null, string? footer = null)
    {
        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle
        {
            Fill = new SvgColourServer(backgroundColor),
            X = 0,
            Y = 0,
            Height = 72,
            Width = 72,
        });

        doc.Children.Add(new SvgText(count.ToString())
        {
            FontSize = 25,
            TextAnchor = SvgTextAnchor.Middle,
            FontWeight = SvgFontWeight.Bold,
            Fill = new SvgColourServer(foregroundColor),
            X = new SvgUnitCollection { new(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new(SvgUnitType.Pixel, 46.5f) },
        });

        if (header != null)
        {
            doc.Children.Add(new SvgText(header)
            {
                FontSize = 15,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Fill = new SvgColourServer(foregroundColor),
                X = new SvgUnitCollection { new(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new(SvgUnitType.Pixel, 15) },
            });
        }

        if (footer != null)
        {
            doc.Children.Add(new SvgText(footer)
            {
                FontSize = 15,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Fill = new SvgColourServer(foregroundColor),
                X = new SvgUnitCollection { new(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new(SvgUnitType.Pixel, 66) },
            });
        }

        if (BadgeImageFilePath != null)
        {
            var imageContent = Convert.ToBase64String(File.ReadAllBytes(BadgeImageFilePath));
            doc.Children.Add(new SvgImage
            {
                Href = $"data:image/png;base64,{imageContent}",
                Width = 25,
                Height = 25,
                X = new SvgUnit(SvgUnitType.Pixel, 46),
                Y = new SvgUnit(SvgUnitType.Pixel, 1),
            });
        }

        using MemoryStream ms = new MemoryStream();
        doc.Write(ms);
        ms.Position = 0;
        using var reader = new StreamReader(ms, System.Text.Encoding.UTF8);

        return reader.ReadToEnd();
    }

    public string CreateNoConnectionSvg(string text, Color textColor, Color backgroundColor)
    {
        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle
        {
            Fill = new SvgColourServer(backgroundColor),
            X = 0,
            Y = 0,
            Height = 72,
            Width = 72,
        });

        doc.Children.Add(new SvgText(text)
        {
            FontSize = 25,
            TextAnchor = SvgTextAnchor.Middle,
            FontWeight = SvgFontWeight.Bold,
            Fill = new SvgColourServer(textColor),
            X = new SvgUnitCollection { new(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new(SvgUnitType.Pixel, 48.5f) },
        });

        if (BadgeImageFilePath != null)
        {
            var imageContent = Convert.ToBase64String(File.ReadAllBytes(BadgeImageFilePath));
            doc.Children.Add(new SvgImage
            {
                Href = $"data:image/png;base64,{imageContent}",
                Width = 25,
                Height = 25,
                X = new SvgUnit(SvgUnitType.Pixel, 46),
                Y = new SvgUnit(SvgUnitType.Pixel, 1),
            });
        }

        using var ms = new MemoryStream();
        doc.Write(ms);
        ms.Position = 0;
        using var reader = new StreamReader(ms, System.Text.Encoding.UTF8);

        return reader.ReadToEnd();
    }
}

class AnimatedIcon
{
    private readonly IconCreator _iconCreator;
    private readonly CancellationTokenSource _cancellationTokenSource = new();

    public AnimatedIcon(string? badgeImageFilePath)
    {
        _iconCreator = new IconCreator(badgeImageFilePath);
    }

    public int Count { get; set; }
    public Color ForegroundColor { get; set; } = Color.White;
    public Color BackgroundColor { get; set; }
    public string? Header { get; set; }
    public string? Footer { get; set; }

    public Func<string, Task> OnIconCreated { get; set; } = content => Task.CompletedTask;

    public async Task AnimateFooter()
    {
        var token = _cancellationTokenSource.Token;

        while (!token.IsCancellationRequested)
        {
            if (Footer != null)
            {
                for (var i = 0; i < Footer.Length - 6; i += System.Globalization.StringInfo.GetNextTextElement(Footer, i).Length)
                {
                    var animatedContent = _iconCreator.CreateNotificationSvg(
                        Count,
                        ForegroundColor,
                        BackgroundColor,
                        header: Header,
                        footer: Footer[i..]);

                    if (token.IsCancellationRequested)
                    {
                        return;
                    }

                    await OnIconCreated(animatedContent);

                    await Task.Delay(200);
                }
            }

            await Task.Delay(1000);

            var content = _iconCreator.CreateNotificationSvg(
                Count,
                ForegroundColor,
                BackgroundColor,
                header: Header,
                footer: Footer);

            if (token.IsCancellationRequested)
            {
                return;
            }

            await OnIconCreated(content);

            await Task.Delay(5000);
        }
    }

    public void CancelAnimation() => _cancellationTokenSource.Cancel();
}

class AnimatedIconLoader : IDisposable
{
    private readonly object _lockObject = new object();
    private bool _isDisposed = false;
    private AnimatedIcon? _animatedIcon;

    public void LoadAndAnimate(AnimatedIcon animatedIcon)
    {
        lock (_lockObject)
        {
            if (_isDisposed)
            {
                return;
            }

            _animatedIcon?.CancelAnimation();

            _animatedIcon = animatedIcon;
            _animatedIcon.AnimateFooter();
        }
    }

    public void CancelAnimation()
    {
        lock (_lockObject)
        {
            _animatedIcon?.CancelAnimation();
        }
    }

    public void Dispose()
    {
        lock (_lockObject)
        {
            _animatedIcon?.CancelAnimation();
            _animatedIcon = null;

            _isDisposed = true;
        }
    }
}