namespace Elgato.Plugins.Microsoft365;

public interface IAction
{
    string? AppId { get; }

    void OnAccountRemoved(string accountId);
}

public class ActionNotifier
{
    public static ActionNotifier Instance = new ActionNotifier();

    private List<IAction> _actions = new List<IAction>();

    public void RegisterAction(IAction action)
    {
        _actions.Add(action);
    }

    public void UnregisterAction(IAction action)
    {
        _actions.Remove(action);
    }

    public void NotifyAllAboutRemovedAccount(string? appId, string accountId) => OnAccountRemoved(appId, accountId);

    private void OnAccountRemoved(string? appId, string accountId)
    {
        var targetActions = _actions.Where(x => x.AppId?.ToLowerInvariant() == appId?.ToLowerInvariant()).ToList();

        targetActions.ForEach(x => x.OnAccountRemoved(accountId));
    }
}