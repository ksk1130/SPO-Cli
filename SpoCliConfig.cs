using System;

/// <summary>
/// 環境変数から必須設定を読み込む。
/// </summary>
internal sealed class SpoCliConfig
{
    public string ClientId { get; }
    public string TenantId { get; }

    private SpoCliConfig(string clientId, string tenantId)
    {
        ClientId = clientId;
        TenantId = tenantId;
    }

    /// <summary>
    /// SPO_CLIENT_ID と SPO_TENANT_ID を読み取る。
    /// </summary>
    public static SpoCliConfig Load()
    {
        var clientId = Environment.GetEnvironmentVariable("SPO_CLIENT_ID");
        if (string.IsNullOrWhiteSpace(clientId))
        {
            throw new InvalidOperationException("Missing SPO_CLIENT_ID. Set it to your Entra ID app client ID.");
        }

        var tenantId = Environment.GetEnvironmentVariable("SPO_TENANT_ID");
        if (string.IsNullOrWhiteSpace(tenantId))
        {
            tenantId = "organizations";
        }

        return new SpoCliConfig(clientId, tenantId);
    }
}
