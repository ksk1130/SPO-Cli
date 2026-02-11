using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft.SharePoint.Client;

/// <summary>
/// MSALのトークン取得とCSOMコンテキスト生成を担う。
/// </summary>
internal sealed class SpoAuth
{
    private readonly IPublicClientApplication _app;

    private SpoAuth(IPublicClientApplication app)
    {
        _app = app;
    }

    /// <summary>
    /// パブリッククライアントを構築し、トークンキャッシュ永続化を有効化する。
    /// </summary>
    public static async Task<SpoAuth> CreateAsync(SpoCliConfig config)
    {
        var app = PublicClientApplicationBuilder
            .Create(config.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, config.TenantId)
            .WithRedirectUri("http://localhost")
            .Build();

        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "spo-cli");
        Directory.CreateDirectory(cacheDir);

        var storageProperties = new StorageCreationPropertiesBuilder("msal.cache", cacheDir)
            .Build();
        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
        cacheHelper.RegisterCache(app.UserTokenCache);

        return new SpoAuth(app);
    }

    /// <summary>
    /// サイトURLのテナントルート向けアクセストークンを取得する。
    /// </summary>
    public async Task<string> AcquireTokenAsync(string siteUrl, bool interactive)
    {
        var scopes = new[] { $"{UrlHelpers.GetTenantRoot(siteUrl)}/.default" };
        var accounts = await _app.GetAccountsAsync();

        try
        {
            var account = accounts.FirstOrDefault();
            if (account == null)
            {
                throw new MsalUiRequiredException("no_account", "No cached account.");
            }

            var result = await _app.AcquireTokenSilent(scopes, account).ExecuteAsync();
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            if (!interactive)
            {
                throw;
            }

            var result = await _app.AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync();
            return result.AccessToken;
        }
    }

    /// <summary>
    /// Bearerトークン付きのCSOMコンテキストを生成する。
    /// </summary>
    public async Task<ClientContext> CreateContextAsync(string siteUrl, bool interactive)
    {
        var token = await AcquireTokenAsync(siteUrl, interactive);
        var context = new ClientContext(siteUrl);
        context.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = $"Bearer {token}";
        };

        return context;
    }
}
