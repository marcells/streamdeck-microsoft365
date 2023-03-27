using BarRaider.SdTools;
using BarRaider.SdTools.Events;
using BarRaider.SdTools.Payloads;
using BarRaider.SdTools.Wrappers;
using Newtonsoft.Json.Linq;

namespace Elgato.Plugins.Microsoft365;

public interface IPluginSettings
{
    abstract static IPluginSettings CreateDefaultSettings();

    string? AppId { get; set; }

    string? Account { get; set; }
}

public abstract class GraphAction<TSettings> : KeyAndEncoderBase, IAction
    where TSettings : IPluginSettings
{
    private GraphAuthenticator? _graphAuthenticator;

    public GraphAction(ISDConnection connection, InitialPayload payload) : base(connection, payload)
    {
        ActionNotifier.Instance.RegisterAction(this);

        connection.OnSendToPlugin += OnSendToPlugin;
        connection.OnPropertyInspectorDidAppear += OnPropertyInspectorDidAppear;

        if (payload.Settings == null || payload.Settings.Count == 0)
        {
            Settings = (TSettings)TSettings.CreateDefaultSettings();
            Connection.SetSettingsAsync(JObject.FromObject(Settings));
        }
        else
        {
            Settings = payload.Settings.ToObject<TSettings>()!;
        }

        InitializePlugin();
    }

    public TSettings Settings { get; set; }

    public bool IsGraphApiInitialized => _graphAuthenticator != null ? _graphAuthenticator.IsInitialized : false;

    public string? AppId => Settings?.AppId;

    public Microsoft.Graph.GraphServiceClient GraphApp =>
        _graphAuthenticator != null && _graphAuthenticator.IsInitialized
            ? _graphAuthenticator.GetApp()
            : throw new InvalidOperationException("GraphAuthenticator not initialized.");

    public override void KeyPressed(KeyPayload payload)
    {
    }

    public override void KeyReleased(KeyPayload payload)
    {
    }

    public override void ReceivedSettings(ReceivedSettingsPayload payload)
    {
        Tools.AutoPopulateSettings(Settings, payload.Settings);

        InitializePlugin();
    }

    public override void ReceivedGlobalSettings(ReceivedGlobalSettingsPayload payload)
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

    public override void Dispose()
    {
        ActionNotifier.Instance.UnregisterAction(this);
    }

    protected abstract Task OnPluginInitialized();

    public async void OnAccountRemoved(string accountId)
    {
        await RemoveAccount(accountId);
    }

    private async Task RemoveSelectedAccount()
    {
        if (string.IsNullOrWhiteSpace(Settings.Account))
            return;

        var accountId = Settings.Account;
        await RemoveAccount(accountId);

        ActionNotifier.Instance.NotifyAllAboutRemovedAccount(Settings.AppId, accountId);
    }

    private async Task RemoveAccount(string? accountId)
    {
        if (string.IsNullOrWhiteSpace(accountId))
            return;

        if (Settings.Account?.ToLowerInvariant() == accountId?.ToLowerInvariant())
        {
            Settings.Account = string.Empty;
            await Connection.SetSettingsAsync(JObject.FromObject(Settings));

            if (_graphAuthenticator != null)
                await _graphAuthenticator.RemoveAccount(accountId);

            await SendAccountsToPropertyInspector();
        }
    }

    private async void InitializePlugin()
    {
        _graphAuthenticator  = new GraphAuthenticator(new GraphSettings { ClientId = Settings?.AppId, AccountId = Settings?.Account });
        await SendAccountsToPropertyInspector();
        
        if (string.IsNullOrWhiteSpace(Settings?.AppId) || string.IsNullOrWhiteSpace(Settings?.Account))
            return;

        await _graphAuthenticator.InitializeAsync();

        await SendAccountsToPropertyInspector();

        await OnPluginInitialized();
    }

    private async void OnSendToPlugin(object? sender, SDEventReceivedEventArgs<SendToPlugin> e)
    {
        var operation = e.Event.Payload.GetValue("operation")?.ToString();

        if (operation == "add")
        {
            var authenticator  = new GraphAuthenticator(new GraphSettings { ClientId = Settings?.AppId, AccountId = null });
            await authenticator.InitializeAsync();

            await SendAccountsToPropertyInspector();
        }
        else if (operation == "remove")
        {
            await RemoveSelectedAccount();
        }
    }

    private async void OnPropertyInspectorDidAppear(object? sender, SDEventReceivedEventArgs<PropertyInspectorDidAppear> e)
    {
        await SendAccountsToPropertyInspector();
    }

    private async Task SendAccountsToPropertyInspector()
    {
        if (_graphAuthenticator == null)
        {
            await SendMessageToPropertyInspector("loadedAccounts", new { });
            return;
        }

        var accounts = await _graphAuthenticator.GetAccounts();
        await Connection.SendToPropertyInspectorAsync(JObject.FromObject(new { message = "accountsLoaded", data = new { accounts = accounts, currentAccount = Settings.Account }}));
    }

    private async Task SendMessageToPropertyInspector(string message, object data)
    {
        await Connection.SendToPropertyInspectorAsync(JObject.FromObject(new { message, data}));
    }
}