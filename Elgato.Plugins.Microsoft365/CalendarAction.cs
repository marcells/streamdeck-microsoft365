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

    [JsonProperty(PropertyName = "account")]
    public string? Account { get; set; }
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
        
        var iconCreator = new IconCreator("Assets\\calendar.png");

        var content = iconCreator.CreateNotificationSvg(result, GetBackgroundColorForEventTimes(nextEventStartsIn, currentEventRunningFor));

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }

    private async Task NoConnectionInfo()
    {
        var iconCreator = new IconCreator("Assets\\calendar.png");

        var content = iconCreator.CreateNoConnectionSvg();

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }
}
