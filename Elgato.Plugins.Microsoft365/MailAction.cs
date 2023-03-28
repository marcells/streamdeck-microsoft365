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
}

[PluginActionId("es.mspi.microsoft.mail")]
public class MailAction : GraphAction<MailPluginSettings>
{
    private AnimatedIcon? _animatedIcon;
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

    public override void Dispose()
    {
        _animatedIcon?.CancelAnimation();

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

    private static Color GetBackgroundColorForNumberOfMails(int numberOfMails) => numberOfMails == 0? Color.LightGray : Color.Yellow;
    
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
        var subject = messages.Select(x => $"{(x.Subject ?? "No subject")} [{(x.From?.EmailAddress?.Name ?? x.From?.EmailAddress?.Address ?? "Unknown")}]").FirstOrDefault();

        return subject;
    }

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
            });

        var numberOfMails = result.GetValueOrDefault();

        var subject = numberOfMails > 0 ? await TryGetSubjectOfLatestMail() : null;

        _animatedIcon?.CancelAnimation();

        _animatedIcon = new AnimatedIcon("Assets\\mail.png")
        {
            Count = numberOfMails,
            BackgroundColor = GetBackgroundColorForNumberOfMails(numberOfMails),
            Footer = subject,
            OnIconCreated = async content => await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}"),
        };

        _animatedIcon.AnimateFooter();
    }

    private async Task NoConnectionInfo()
    {
        _animatedIcon?.CancelAnimation();

        var iconCreator = new IconCreator("Assets\\mail.png");

        var content = iconCreator.CreateNoConnectionSvg();

        await Connection.SetImageAsync($"data:image/svg+xml;charset=utf8,{content}");
    }
}
