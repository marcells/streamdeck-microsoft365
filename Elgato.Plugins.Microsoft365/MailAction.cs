using System.Diagnostics;
using System.Drawing;
using BarRaider.SdTools;
using Newtonsoft.Json;

namespace Elgato.Plugins.Microsoft365;

public class MailPluginSettings : IPluginSettings
{
    public static IPluginSettings CreateDefaultSettings() => new MailPluginSettings();

    [JsonProperty(PropertyName = "appId")]
    public string? AppId { get; set; }

    [JsonProperty(PropertyName = "account")]
    public string? Account { get; set; }

    [JsonProperty(PropertyName = "unreadColor")]
    public string UnreadColor { get; set; } = "#f1fa8c";

    [JsonProperty(PropertyName = "unreadTextColor")]
    public string UnreadTextColor { get; set; } = "#44475a";

    [JsonProperty(PropertyName = "readColor")]
    public string ReadColor { get; set; } = "#282a36";

    [JsonProperty(PropertyName = "readTextColor")]
    public string ReadTextColor { get; set; } = "#F8F8F2";

    [JsonProperty(PropertyName = "noConnectionColor")]
    public string NoConnectionColor { get; set; } = "#ff5555";

    [JsonProperty(PropertyName = "noConnectionTextColor")]
    public string NoConnectionTextColor { get; set; } = "#44475a";

    [JsonProperty(PropertyName = "showBadge")]
    public bool ShowBadge { get; set; } = true;

    [JsonProperty(PropertyName = "openApp")]
    public bool OpenApp { get; set; }

    [JsonProperty(PropertyName = "appPath")]
    public string? AppPath { get; set; }

    [JsonProperty(PropertyName = "showTitle")]
    public bool ShowTitle { get; set; } = true;
}

[PluginActionId("es.mspi.elgato.plugins.microsoft.mail")]
public class MailAction : GraphAction<MailPluginSettings>
{
    private readonly AnimatedIconLoader _animatedIconLoader = new AnimatedIconLoader();
    private DateTime _lastCheck = DateTime.Now.AddDays(-1);

    public MailAction(ISDConnection connection, InitialPayload payload)
        : base(connection, payload)
    {
    }

    protected override async Task OnPluginInitialized()
    {
        await TryUpdateBadge(true);
    }

    public override void ReceivedSettings(ReceivedSettingsPayload payload)
    {
        base.ReceivedSettings(payload);

        Logger.Instance.LogMessage(TracingLevel.INFO, $"Received settings: {payload.Settings}");

        TryUpdateBadge(true);
    }

    public override async void KeyPressed(KeyPayload payload)
    {
        if (!IsGraphApiInitialized)
        {
            return;
        }

        var result = await GraphApp
            .Me
            .MailFolders["Inbox"]
            .GetAsync();

        if (result != null)
        {
            if (!string.IsNullOrEmpty(Settings.AppPath) && Settings.OpenApp)
            {
                Process.Start(new ProcessStartInfo { FileName = Settings.AppPath, UseShellExecute = true });
            }
            else
            {
                Process.Start(new ProcessStartInfo
                    { FileName = $"https://outlook.live.com/mail/{result.Id}", UseShellExecute = true });
            }
        }

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

    private Color GetBackgroundColorForNumberOfMails(int numberOfMails)
    {
        return numberOfMails == 0
            ? ColorTranslator.FromHtml(Settings.ReadColor)
            : ColorTranslator.FromHtml(Settings.UnreadColor);
    }

    private Color GetForegroundColorForNumberOfMails(int numberOfMails)
    {
        return numberOfMails == 0
            ? ColorTranslator.FromHtml(Settings.ReadTextColor)
            : ColorTranslator.FromHtml(Settings.UnreadTextColor);
    }

    private async Task<string?> TryGetSubjectOfLatestMail()
    {
        var mails = await GraphApp
            .Me
            .MailFolders["Inbox"]
            .Messages
            .GetAsync(x =>
            {
                x.QueryParameters.Filter = "isread eq false";
                x.QueryParameters.Orderby = new[] { "receiveddatetime DESC" };
            });

        var messages = mails?.Value ?? new List<Microsoft.Graph.Models.Message>();
        var subject = messages.Select(x =>
                $"{(x.Subject ?? "No subject")} [{(x.From?.EmailAddress?.Name ?? x.From?.EmailAddress?.Address ?? "Unknown")}]")
            .FirstOrDefault();

        return subject;
    }

    private async Task UpdateBadge()
    {
        Logger.Instance.LogMessage(TracingLevel.INFO, "Updating badge");

        var result = await GraphApp
            .Me
            .MailFolders["Inbox"]
            .Messages
            .Count
            .GetAsync(x =>
            {
                x.QueryParameters.Filter = "isread eq false";
            });

        var numberOfMails = result.GetValueOrDefault();

        var subject = numberOfMails > 0 ? await TryGetSubjectOfLatestMail() : null;
        var badge = Settings.ShowBadge ? "Assets\\mail.png" : null;
        var icon = new AnimatedIcon(badge)
        {
            Count = numberOfMails,
            ForegroundColor = GetForegroundColorForNumberOfMails(numberOfMails),
            BackgroundColor = GetBackgroundColorForNumberOfMails(numberOfMails),
            Footer = Settings.ShowTitle ? subject : null,
            OnIconCreated = async content => await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}"),
        };

        _animatedIconLoader.LoadAndAnimate(icon);
    }

    private async Task NoConnectionInfo()
    {
        _animatedIconLoader.CancelAnimation();

        var badge = Settings.ShowBadge ? "Assets\\mail.png" : null;
        var iconCreator = new IconCreator(badge);

        var content =
            iconCreator.CreateNoConnectionSvg("Conn", ColorTranslator.FromHtml(Settings.NoConnectionTextColor), ColorTranslator.FromHtml(Settings.NoConnectionColor));

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }
}