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
        // localhost の場合は認証をスキップ（モックサーバー用）
        if (IsLocalhost(siteUrl))
        {
            return "mock-token-for-localhost";
        }

        var result = await AcquireTokenResultAsync(siteUrl, interactive);
        return result.AccessToken;
    }

    /// <summary>
    /// トークン取得結果（ExpiresOn含む）を返す。
    /// </summary>
    public async Task<AuthenticationResult> AcquireTokenResultAsync(
        string siteUrl,
        bool interactive,
        Prompt? prompt = null)
    {
        // localhost の場合はダミーの AuthenticationResult を返す（モックサーバー用）
        // 注: AuthenticationResult は sealed なので実際には使用されない想定
        // CreateContextAsync で localhost チェックを行いトークン取得をスキップする
        if (IsLocalhost(siteUrl))
        {
            // ダミーを返すが、実際には CreateContextAsync で使用されない
            throw new NotSupportedException("Localhost does not require authentication. This should not be called.");
        }

        var scopes = new[] { UrlHelpers.GetTenantRoot(siteUrl) + "/.default" };
        var accounts = await _app.GetAccountsAsync();

        try
        {
            var account = accounts.FirstOrDefault();
            if (account == null)
            {
                throw new MsalUiRequiredException("no_account", "No cached account.");
            }

            return await _app.AcquireTokenSilent(scopes, account).ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            if (!interactive)
            {
                throw;
            }
        }

        var effectivePrompt = prompt ?? Prompt.SelectAccount;
        return await _app.AcquireTokenInteractive(scopes)
            .WithPrompt(effectivePrompt)
            .ExecuteAsync();
    }

    /// <summary>
    /// Bearerトークン付きのCSOMコンテキストを生成する。
    /// </summary>
    public async Task<ClientContext> CreateContextAsync(
        string siteUrl,
        bool interactive,
        bool showExpiresOn = false,
        AuthenticationResult tokenResult = null)
    {
        string token;
        
        // localhost の場合は認証をスキップ
        if (IsLocalhost(siteUrl))
        {
            token = "mock-token-for-localhost";
            if (showExpiresOn)
            {
                Console.WriteLine("Using mock authentication for localhost");
            }
        }
        else
        {
            var result = tokenResult ?? await AcquireTokenResultAsync(siteUrl, interactive);
            if (showExpiresOn)
            {
                DisplayExpiresOn(result);
            }
            token = result.AccessToken;
        }

        var context = new ClientContext(siteUrl);
        context.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
        };

        return context;
    }

    private static void DisplayExpiresOn(AuthenticationResult result)
    {
        var local = result.ExpiresOn.LocalDateTime;
        Console.WriteLine(string.Format("Access token ExpiresOn (local): {0:yyyy-MM-dd HH:mm:ss}", local));
    }

    /// <summary>
    /// URL が localhost を指しているかチェックする。
    /// </summary>
    private static bool IsLocalhost(string url)
    {
        Uri uri;
        if (Uri.TryCreate(url, UriKind.Absolute, out uri))
        {
            return uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                || uri.Host.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }
}
