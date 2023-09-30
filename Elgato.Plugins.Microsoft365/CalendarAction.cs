using System.Diagnostics;
using System.Drawing;
using BarRaider.SdTools;
using Newtonsoft.Json;

namespace Elgato.Plugins.Microsoft365;

public class CalendarPluginSettings : IPluginSettings
{
    public static IPluginSettings CreateDefaultSettings() => new CalendarPluginSettings();

    [JsonProperty(PropertyName = "appId")]
    public string? AppId { get; set; }

    [JsonProperty(PropertyName = "account")]
    public string? Account { get; set; }
}

[PluginActionId("es.mspi.elgato.plugins.microsoft.calendar")]
public class CalendarAction : GraphAction<CalendarPluginSettings>
{
    private readonly AnimatedIconLoader _animatedIconLoader = new AnimatedIconLoader();
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
        {
            return;
        }

        Process.Start(new ProcessStartInfo { FileName = $"https://outlook.live.com/calendar", UseShellExecute = true });

        await TryUpdateBadge(true);
    }

    public override void Dispose()
    {
        _animatedIconLoader.Dispose();

        base.Dispose();
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

        var today = await TryGetEventsForDateRange(DateTime.Now.ToUniversalTime(), DateTime.Now.Date.AddDays(1).AddMinutes(-1).ToUniversalTime());
        
        if (today.Any())
        {
            return today;
        }

        var tomorrow = await TryGetEventsForDateRange(DateTime.Now.Date.AddDays(1).ToUniversalTime(), DateTime.Now.Date.AddDays(2).AddMinutes(-1).ToUniversalTime());

        return tomorrow;
    }

    private static Color GetBackgroundColorForEventTimes(TimeSpan? nextEventStartsIn, TimeSpan? currentEventRunningFor)
    {
        if (currentEventRunningFor?.TotalMinutes < 3)
        {
            return Color.Red;
        }

        if (nextEventStartsIn?.TotalMinutes <= 3)
        {
            return Color.Red;
        }

        if (nextEventStartsIn?.TotalMinutes <= 15)
        {
            return Color.OrangeRed;
        }

        if (!currentEventRunningFor.HasValue && !nextEventStartsIn.HasValue)
        {
            return Color.LightGray;
        }

        return Color.Yellow;
    }

    private Color GetForegroundColorForEventTimes(TimeSpan? nextEventStartsIn, TimeSpan? currentEventRunningFor)
    {
        // TODO: Implement Settings :)
        return Color.White;
    }

    private async Task UpdateBadge()
    {
        var results = await TryGetEventsForTodayOrTomorrow();

        var result = results.Count;

        static DateTime ParseDateTimeByEvent(Microsoft.Graph.Models.Event @event, bool isEndDate)
        {
            var dateTimeTimeZone = isEndDate ? @event.End! : @event.Start!;
            var isAllDayEvent = @event.IsAllDay.GetValueOrDefault();

            var date = isAllDayEvent
                ? DateTime.Parse(dateTimeTimeZone!.DateTime!).Date
                : DateTime.Parse(dateTimeTimeZone!.DateTime!).ToLocalTime();

            return isEndDate && isAllDayEvent ? date.AddDays(1).AddMinutes(-1) : date;
        }

        var times = results
            .Select(x => new 
            { 
                StartTime = ParseDateTimeByEvent(x, false),
                EndTime = ParseDateTimeByEvent(x, true),
                Subject = x.Subject,
            })
            .OrderBy(x => x.StartTime)
            .ToList();

        var currentEvent = times.FirstOrDefault(x => x.StartTime < DateTime.Now && DateTime.Now < x.EndTime);
        var nextEvent = times.FirstOrDefault(x => x.StartTime > DateTime.Now);

        var nextEventStartsIn = nextEvent?.StartTime != null
            ? nextEvent.StartTime - DateTime.Now
            : (TimeSpan?)null;
        
        var currentEventRunningFor = currentEvent?.StartTime != null
            ? DateTime.Now - currentEvent.StartTime
            : (TimeSpan?)null;
        
        _animatedIconLoader.LoadAndAnimate(new AnimatedIcon("Assets\\calendar.png")
        {
            Count = result,
            ForegroundColor = GetForegroundColorForEventTimes(nextEventStartsIn, currentEventRunningFor),
            BackgroundColor = GetBackgroundColorForEventTimes(nextEventStartsIn, currentEventRunningFor),
            Header = nextEvent?.StartTime.ToShortTimeString(),
            Footer = nextEvent?.Subject,
            OnIconCreated = async content => await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}"),
        });
    }

    private async Task NoConnectionInfo()
    {
        _animatedIconLoader.CancelAnimation();

        var iconCreator = new IconCreator("Assets\\calendar.png");

        var content = iconCreator.CreateNoConnectionSvg("Conn", ColorTranslator.FromHtml("#F8F8F2"), ColorTranslator.FromHtml("#FF5555"));

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }
}
