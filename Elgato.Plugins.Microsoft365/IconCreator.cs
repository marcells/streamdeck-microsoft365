using System.Drawing;
using Svg;

namespace Elgato.Plugins.Microsoft365;

class IconCreator
{
    public IconCreator(string badgeImageFilePath)
    {
        BadgeImageFilePath = badgeImageFilePath;        
    }

    public string BadgeImageFilePath { get; init; }

    public string CreateNotificationSvg(int count, Color backgroundColor, string? header = null, string? footer = null)
    {
        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle()
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
            Color = new SvgColourServer(Color.Blue),
            X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 48.5f) },
        });

        if (header != null)
        {
            doc.Children.Add(new SvgText(header)
            {
                FontSize = 15,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Color = new SvgColourServer(Color.Blue),
                X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 15) },
            });
        }

        if (footer != null)
        {
            doc.Children.Add(new SvgText(footer)
            {
                FontSize = 15,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Color = new SvgColourServer(Color.Blue),
                X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 64) },
            });
        }

        if (BadgeImageFilePath != null) 
        {
            var imageContent = Convert.ToBase64String(File.ReadAllBytes(BadgeImageFilePath));
            doc.Children.Add(new SvgImage() {
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

    public string CreateNoConnectionSvg()
    {
        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle()
        {
            Fill = new SvgColourServer(Color.LightGray),
            X = 0,
            Y = 0,
            Height = 72,
            Width = 72,
        });

        doc.Children.Add(new SvgText(("Nope"))
        {
            FontSize = 25,
            TextAnchor = SvgTextAnchor.Middle,
            FontWeight = SvgFontWeight.Bold,
            Color = new SvgColourServer(Color.Blue),
            X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 48.5f) },
        });

        var imageContent = Convert.ToBase64String(File.ReadAllBytes(BadgeImageFilePath));
        doc.Children.Add(new SvgImage() {
            Href = $"data:image/png;base64,{imageContent}",
            Width = 25,
            Height = 25,
            X = new SvgUnit(SvgUnitType.Pixel, 46),
            Y = new SvgUnit(SvgUnitType.Pixel, 1),
        });

        using MemoryStream ms = new MemoryStream();
        doc.Write(ms);
        ms.Position = 0;
        using var reader = new StreamReader(ms, System.Text.Encoding.UTF8);
        
        return reader.ReadToEnd();
    }
}

class AnimatedIcon
{
    private IconCreator _iconCreator;
    private CancellationTokenSource _cancellationTokenSource = new CancellationTokenSource();

    public AnimatedIcon(string badgeImageFilePath)
    {
        _iconCreator = new IconCreator(badgeImageFilePath);
    }

    public int Count { get; set; }
    public Color BackgroundColor { get; set; }
    public string? Header { get; set; }
    public string? Footer { get; set; }

    public Func<string, Task> OnIconCreated { get; set; } = content => Task.CompletedTask;

    public async void AnimateFooter()
    {
        _cancellationTokenSource = new CancellationTokenSource();
        var token = _cancellationTokenSource.Token;

        for(;;)
        {
            if (Footer != null)
            {
                for (var i = 0; i < Footer.Length; i++)
                {
                    var animatedContent = _iconCreator.CreateNotificationSvg(
                                            Count,
                                            BackgroundColor,
                                            header: Header,
                                            footer: Footer.Substring(i));

                    if (token.IsCancellationRequested)
                        return;

                    await OnIconCreated(animatedContent);

                    await Task.Delay(200);
                }
            }

            await Task.Delay(1000);

            var content = _iconCreator.CreateNotificationSvg(
                            Count,
                            BackgroundColor,
                            header: Header,
                            footer: Footer);

            if (token.IsCancellationRequested)
                return;

            await OnIconCreated(content);

            await Task.Delay(5000);
        }
    }

    public void CancelAnimation() => _cancellationTokenSource.Cancel();
}