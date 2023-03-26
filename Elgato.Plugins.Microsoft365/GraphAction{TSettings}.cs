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

public interface IPluginSettings
{
    abstract static IPluginSettings CreateDefaultSettings();

    string AppId { get; set; }
}

public abstract class GraphAction<TSettings> : KeyAndEncoderBase
    where TSettings : IPluginSettings
{
    private GraphAuthenticator _graphAuthenticator;

    public GraphAction(ISDConnection connection, InitialPayload payload) : base(connection, payload)
    {
        connection.OnSendToPlugin += SendToPlugin;
        
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

    public bool IsGraphApiInitialized => _graphAuthenticator.IsInitialized;

    public Microsoft.Graph.GraphServiceClient GraphApp => _graphAuthenticator.GetApp();

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

    protected async Task ResetPlugin()
    {
        Settings = (TSettings)TSettings.CreateDefaultSettings();
        await Connection.SetSettingsAsync(JObject.FromObject(Settings));

        await _graphAuthenticator.Reset();
    }

    protected async void InitializePlugin()
    {
        _graphAuthenticator  = new GraphAuthenticator(new GraphSettings { ClientId = Settings.AppId });
        await _graphAuthenticator.InitializeAsync();

        await OnPluginInitialized();
    }

    protected abstract Task OnPluginInitialized();

    private async void SendToPlugin(object? sender, SDEventReceivedEventArgs<SendToPlugin> e)
    {
        var operation = e.Event.Payload.GetValue("operation")?.ToString();

        if (operation == "clear")
        {
            await ResetPlugin();
        }
    }
}