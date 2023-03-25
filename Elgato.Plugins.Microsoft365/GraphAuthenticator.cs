using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft.Kiota.Abstractions.Authentication;

namespace Elgato.Plugins.Microsoft365;

public class GraphAuthenticator
{
    private GraphSettings _graphSettings;
    private GraphServiceClient? _userClient;
        
    public GraphAuthenticator(GraphSettings graphSettings)
    {
        _graphSettings = graphSettings;
    }

    public bool IsInitialized { get; private set; }

    private async Task<GraphServiceClient> InitializeGraphForUserAuth()
    {
        var (app, cacheHelper) = await GetAppAndCacheHelper();

        var tokenProvider = new TokenProvider(app, AppConfig.Scopes);
        var accessTokenProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
        
        return new GraphServiceClient(accessTokenProvider);
    }

    private async Task<(IPublicClientApplication App, MsalCacheHelper CacheHelper)> GetAppAndCacheHelper()
    {
        var storageProperties = new StorageCreationPropertiesBuilder(TokenCacheConfig.CacheFileName, TokenCacheConfig.CacheDir)
            .WithLinuxKeyring(
                TokenCacheConfig.LinuxKeyRingSchema,
                TokenCacheConfig.LinuxKeyRingCollection,
                TokenCacheConfig.LinuxKeyRingLabel,
                TokenCacheConfig.LinuxKeyRingAttr1,
                TokenCacheConfig.LinuxKeyRingAttr2)
            .WithMacKeyChain(
                TokenCacheConfig.KeyChainServiceName,
                TokenCacheConfig.KeyChainAccountName)
            .Build();

        var builder = PublicClientApplicationBuilder.Create(_graphSettings.ClientId)
                .WithAuthority($"https://login.microsoftonline.com/{AppConfig.TenantId}")
                .WithDefaultRedirectUri();

        var app = builder.Build();

        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties );
        cacheHelper.RegisterCache(app.UserTokenCache);

        return (app, cacheHelper);
    }

    public async Task InitializeAsync()
    {
        if (IsInitialized || string.IsNullOrWhiteSpace(_graphSettings.ClientId))
            return;

        try
        {
            _userClient = await InitializeGraphForUserAuth();

            var me = await _userClient.Me.GetAsync();

            IsInitialized = true;
        }
        catch
        {
        }
    }

    public async Task Reset()
    {
        if (string.IsNullOrWhiteSpace(_graphSettings.ClientId))
            return;

        var (app, _) = await GetAppAndCacheHelper();

        var accounts = await app.GetAccountsAsync();
        accounts.ToList().ForEach(async account => await app.RemoveAsync(account));

        IsInitialized = false;
    }

    public GraphServiceClient GetApp() 
    {
        if (_userClient == null)
            throw new InvalidOperationException();
        
        return _userClient;
    }
}

public class TokenProvider : IAccessTokenProvider
{
    private IPublicClientApplication _app;
    private string[] _scopes;

    public TokenProvider(IPublicClientApplication app, string[] scopes)
    {
        _app = app;
        _scopes = scopes;

        AllowedHostsValidator = new AllowedHostsValidator();
    }

    public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = default, CancellationToken cancellationToken = default)
    {
        AuthenticationResult result;

        try
        {
            var accounts = await _app.GetAccountsAsync();
            var account = accounts.SingleOrDefault();

            result = await _app.AcquireTokenSilent(_scopes, account).ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            result = await _app.AcquireTokenInteractive(_scopes)
                        .WithUseEmbeddedWebView(false)
                        .ExecuteAsync();
        }

        return result.AccessToken;
    }

    public AllowedHostsValidator AllowedHostsValidator { get; }
}

public class GraphSettings
{
    public string ClientId { get; set; } // db75ff81-2cc6-4048-ab2a-9092cf25e7f6
}

static class AppConfig
{
    public const string TenantId = "common";
    public readonly static string[] Scopes = new string[] { "offline_access", "user.read", "mail.read", "calendars.read" };
}

static class TokenCacheConfig
{
    public const string CacheFileName = "myapp_msal_cache.txt";
    public readonly static string CacheDir = MsalCacheHelper.UserRootDirectory;

    public const string KeyChainServiceName = "myapp_msal_service";
    public const string KeyChainAccountName = "myapp_msal_account";

    public const string LinuxKeyRingSchema = "com.contoso.devtools.tokencache";
    public const string LinuxKeyRingCollection = MsalCacheHelper.LinuxKeyRingDefaultCollection;
    public const string LinuxKeyRingLabel = "MSAL token cache for all Contoso dev tool apps.";
    public static readonly KeyValuePair<string, string> LinuxKeyRingAttr1 = new KeyValuePair<string, string>("Version", "1");
    public static readonly KeyValuePair<string, string> LinuxKeyRingAttr2 = new KeyValuePair<string, string>("ProductGroup", "MyApps");
}