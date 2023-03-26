using System.Diagnostics;
using System.Drawing;
using BarRaider.SdTools;
using Newtonsoft.Json;
using Svg;

namespace Elgato.Plugins.Microsoft365;

public class CalendarPluginSettings : IPluginSettings
{
    public static IPluginSettings CreateDefaultSettings() => new CalendarPluginSettings();

    [JsonProperty(PropertyName = "appId")]
    public string? AppId { get; set; }
}

[PluginActionId("es.mspi.microsoft.calendar")]
public class CalendarAction : GraphAction<CalendarPluginSettings>
{
    private DateTime _lastCheck = DateTime.Now.AddDays(-1);

    public CalendarAction(ISDConnection connection, InitialPayload payload)
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

        Process.Start(new ProcessStartInfo { FileName = $"https://outlook.live.com/calendar", UseShellExecute = true });

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

    private async Task<IReadOnlyList<Microsoft.Graph.Models.Event>> TryGetEventsForTodayOrTomorrow()
    {
        async Task<IReadOnlyList<Microsoft.Graph.Models.Event>> TryGetEventsForDateRange(DateTime from, DateTime to)
        {
            var fromIso = from.ToString("s");
            var toIso = to.ToString("s");

            var events = await GraphApp
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
