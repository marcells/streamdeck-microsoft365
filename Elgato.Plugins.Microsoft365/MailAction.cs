using System.Diagnostics;
using System.Drawing;
using BarRaider.SdTools;
using Newtonsoft.Json;
using Svg;

namespace Elgato.Plugins.Microsoft365;

public class MailPluginSettings : IPluginSettings
{
    public static IPluginSettings CreateDefaultSettings() => new MailPluginSettings();

    [JsonProperty(PropertyName = "appId")]
    public string? AppId { get; set; }
}

[PluginActionId("es.mspi.microsoft.mail")]
public class MailAction : GraphAction<MailPluginSettings>
{
    private DateTime _lastCheck = DateTime.Now.AddDays(-1);

    public MailAction(ISDConnection connection, InitialPayload payload)
        : base(connection, payload)
    {
    }

    protected override async Task OnPluginInitialized()
    {
        await TryUpdateBadge(true);
    }

    public override async void KeyPressed(KeyPayload payload)
    {
        if (!IsGraphApiInitialized)
            return;

        var result = await GraphApp
            .Me
            .MailFolders["Inbox"]
            .GetAsync();

        if (result != null)
            Process.Start(new ProcessStartInfo { FileName = $"https://outlook.live.com/mail/{result.Id}", UseShellExecute = true });

        await TryUpdateBadge(true);
    }

    public override async void OnTick()
    {
        await TryUpdateBadge(false);
    }

    private async Task TryUpdateBadge(bool forceUpdate)
    {
        if (!IsGraphApiInitialized)
        {
            await NoConnectionInfo();
            return;
        }

        var diff = DateTime.Now - _lastCheck;

        if (forceUpdate || diff.TotalMinutes > 2.0)
        {
            _lastCheck = DateTime.Now;
            await UpdateBadge();
        }
    }

    private static Color GetBackgroundColorForNumberOfMails(int numberOfMails) => numberOfMails == 0? Color.LightGray : Color.Yellow;
    
    private async Task UpdateBadge()
    {
        var result = await GraphApp
            .Me
            .MailFolders["Inbox"]
            .Messages
            .Count
            .GetAsync(x =>
            {
                x.QueryParameters.Filter = "isread eq false";
                // x.QueryParameters.Select = new string[] { "From", "IsRead", "ReceivedDateTime", "Subject" };
                // x.QueryParameters.Count = true;
                // x.QueryParameters.Top = 5;
                // x.QueryParameters.Orderby = new string[] { "ReceivedDateTime DESC" };
            });

        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle()
        {
            Fill = new SvgColourServer(GetBackgroundColorForNumberOfMails(result ?? 0)),
            X = 0,
            Y = 0,
            Height = 72,
            Width = 72,
        });

        doc.Children.Add(new SvgText((result ?? 0).ToString())
        {
            FontSize = 25,
            TextAnchor = SvgTextAnchor.Middle,
            FontWeight = SvgFontWeight.Bold,
            Color = new SvgColourServer(Color.Blue),
            X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 48.5f) },
        });

        var imageContent = Convert.ToBase64String(File.ReadAllBytes("Assets\\mail.png"));
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
        var content = reader.ReadToEnd();

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }

    private async Task NoConnectionInfo()
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

        var imageContent = Convert.ToBase64String(File.ReadAllBytes("Assets\\mail.png"));
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
        var content = reader.ReadToEnd();

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }
}
