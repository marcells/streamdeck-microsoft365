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

    public string CreateNotificationSvg(int count, Color backgroundColor)
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