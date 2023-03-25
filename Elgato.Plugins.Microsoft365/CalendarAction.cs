using System.Diagnostics;
using System.Drawing;
using BarRaider.SdTools;
using BarRaider.SdTools.Events;
using BarRaider.SdTools.Payloads;
using BarRaider.SdTools.Wrappers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Svg;

namespace Elgato.Plugins.Microsoft365;

[PluginActionId("es.mspi.microsoft.calendar")]
public class CalendarAction : KeyAndEncoderBase
{
    private PluginSettings _settings;
    private GraphAuthenticator _graphAuthenticator;

    private DateTime _lastCheck = DateTime.Now.AddDays(-1);

    private class PluginSettings
    {
        public static PluginSettings CreateDefaultSettings() => new PluginSettings();

        [JsonProperty(PropertyName = "appId")]
        public string AppId { get; set; }
    }

    public CalendarAction(ISDConnection connection, InitialPayload payload) : base(connection, payload)
    {
        connection.OnSendToPlugin += SendToPlugin;

        if (payload.Settings == null || payload.Settings.Count == 0)
        {
            _settings = PluginSettings.CreateDefaultSettings();
            Connection.SetSettingsAsync(JObject.FromObject(_settings));
        }
        else
        {
            _settings = payload.Settings.ToObject<PluginSettings>()!;
        }

        InitializePlugin();
    }

    private async void SendToPlugin(object? sender, SDEventReceivedEventArgs<SendToPlugin> e)
    {
        var operation = e.Event.Payload.GetValue("operation")?.ToString();

        if (operation == "clear")
        {
            await ResetPlugin();
        }
    }

    private async Task ResetPlugin()
    {
        _settings = PluginSettings.CreateDefaultSettings();
        await Connection.SetSettingsAsync(JObject.FromObject(_settings));

        await _graphAuthenticator.Reset();
    }

    private async void InitializePlugin()
    {
        _graphAuthenticator  = new GraphAuthenticator(new GraphSettings { ClientId = _settings.AppId });
        await _graphAuthenticator.InitializeAsync();

        await TryUpdateBadge(true);
    }

    public async override void ReceivedSettings(ReceivedSettingsPayload payload)
    {
        Tools.AutoPopulateSettings(_settings, payload.Settings);

        await _graphAuthenticator.Reset();
        InitializePlugin();
    }

    public override void ReceivedGlobalSettings(ReceivedGlobalSettingsPayload payload)
    {
    }

    public override async void KeyPressed(KeyPayload payload)
    {
        if (!_graphAuthenticator.IsInitialized)
            return;

        Process.Start(new ProcessStartInfo { FileName = $"https://outlook.live.com/calendar", UseShellExecute = true });

        await TryUpdateBadge(true);
    }

    public override void KeyReleased(KeyPayload payload)
    {
    }

    public override async void OnTick()
    {
        await TryUpdateBadge(false);
    }

    public override void Dispose()
    {
    }

    public override void DialRotate(DialRotatePayload payload)
    {
    }

    public override void DialPress(DialPressPayload payload)
    {
    }

    public override void TouchPress(TouchpadPressPayload payload)
    {
    }

    private async Task TryUpdateBadge(bool forceUpdate)
    {
        if (!_graphAuthenticator.IsInitialized)
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

    private async Task<IReadOnlyList<Microsoft.Graph.Models.Event>> TryGetEventsForTodayOrTomorrow()
    {
        async Task<IReadOnlyList<Microsoft.Graph.Models.Event>> TryGetEventsForDateRange(DateTime from, DateTime to)
        {
            var fromIso = from.ToString("s");
            var toIso = to.ToString("s");

            var events = await _graphAuthenticator
                .GetApp()
                .Me
                .CalendarView
                .GetAsync(x =>
                {
                    x.QueryParameters.StartDateTime = fromIso;
                    x.QueryParameters.EndDateTime = toIso;
                });

            return events?.Value != null ? events.Value : Enumerable.Empty<Microsoft.Graph.Models.Event>().ToList();
        }

        var today = await TryGetEventsForDateRange(DateTime.UtcNow, DateTime.UtcNow.Date.AddDays(1));
        
        if (today.Any())
            return today;

        var tomorrow = await TryGetEventsForDateRange(DateTime.UtcNow.Date.AddDays(1), DateTime.UtcNow.Date.AddDays(2));

        return tomorrow;
    }

    private static Color GetBackgroundColorForEventTimes(TimeSpan? nextEventStartsIn, TimeSpan? currentEventRunningFor)
    {
        if (currentEventRunningFor?.TotalMinutes < 3)
            return Color.Red;

        if (nextEventStartsIn?.TotalMinutes <= 3)
            return Color.Red;

        if (nextEventStartsIn?.TotalMinutes <= 15)
            return Color.OrangeRed;
        
        if (!currentEventRunningFor.HasValue && !nextEventStartsIn.HasValue)
            return Color.LightGray;

        return Color.Yellow;
    }

    private async Task UpdateBadge()
    {
        var results = await TryGetEventsForTodayOrTomorrow();

        var result = results.Count;

        var times = results
            .Select(x => new { StartTime = DateTime.Parse(x.Start!.DateTime!).ToLocalTime(), EndTime = DateTime.Parse(x.End!.DateTime!).ToLocalTime() })
            .OrderBy(x => x.StartTime)
            .ToList();

        var currentEvent = times.Where(x => x.StartTime < DateTime.Now && DateTime.Now < x.EndTime).FirstOrDefault();
        var nextEvent = times.Where(x => x.StartTime > DateTime.Now).FirstOrDefault();

        var nextEventStartsIn = nextEvent?.StartTime != null
            ? nextEvent.StartTime - DateTime.Now
            : (TimeSpan?)null;
        
        var currentEventRunningFor = currentEvent?.StartTime != null
            ? DateTime.Now - currentEvent.StartTime
            : (TimeSpan?)null;
        
        var doc = new SvgDocument
        {
            Width = 72,
            Height = 72,
            ViewBox = new SvgViewBox(0, 0, 72, 72),
        };

        doc.Children.Add(new SvgRectangle()
        {
            Fill = new SvgColourServer(GetBackgroundColorForEventTimes(nextEventStartsIn, currentEventRunningFor)),
            X = 0,
            Y = 0,
            Height = 72,
            Width = 72,
        });

        doc.Children.Add(new SvgText(result.ToString())
        {
            FontSize = 25,
            TextAnchor = SvgTextAnchor.Middle,
            FontWeight = SvgFontWeight.Bold,
            Color = new SvgColourServer(Color.Blue),
            X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 36) },
            Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 48.5f) },
        });

        if (nextEvent != null)
        {
            doc.Children.Add(new SvgText(nextEvent.StartTime.ToShortTimeString())
            {
                FontSize = 15,
                TextAnchor = SvgTextAnchor.Start,
                FontWeight = SvgFontWeight.Normal,
                Color = new SvgColourServer(Color.Blue),
                X = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 2) },
                Y = new SvgUnitCollection { new SvgUnit(SvgUnitType.Pixel, 15) },
            });
        }

        var imageContent = Convert.ToBase64String(File.ReadAllBytes("Assets\\calendar.png"));
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

        var imageContent = Convert.ToBase64String(File.ReadAllBytes("Assets\\calendar.png"));
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
