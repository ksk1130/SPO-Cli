using System;
using System.IO;

/// <summary>
/// SPOとローカル判定向けのURL補助関数。
/// </summary>
internal static class UrlHelpers
{
	/// <summary>
	/// spo:// プレフィックスをデフォルトルートで展開する。
	/// Shared Documents 配下がデフォルトになる。
	/// </summary>
	public static string ExpandSpoShorthand(string value)
	{
		if (!value.StartsWith("spo://", StringComparison.OrdinalIgnoreCase))
		{
			return value;
		}

		var settings = SpoCliSettings.Load();
		if (string.IsNullOrWhiteSpace(settings.DefaultRoot))
		{
			throw new InvalidOperationException("spo:// を使うには先に login でサイトを指定してください。");
		}

		var relative = value.Substring(6);
		if (string.IsNullOrWhiteSpace(relative))
		{
			return CombineUrl(settings.DefaultRoot, "Shared Documents");
		}

		return CombineUrl(settings.DefaultRoot, "Shared Documents", relative);
	}

	/// <summary>
	/// SharePoint Online URLかどうかを判定する。
	/// </summary>
	public static bool IsSpoUrl(string value)
	{
		return TryParseUri(value, out var uri)
			&& (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp)
			&& uri.Host.Contains("sharepoint.com", StringComparison.OrdinalIgnoreCase);
	}

    /// <summary>
    /// サイトURLのスキームとホストを返す。
    /// </summary>
    public static string GetTenantRoot(string url)
    {
        var uri = ParseUri(url);
        return $"{uri.Scheme}://{uri.Host}";
    }

    /// <summary>
    /// 絶対URLをサーバー相対パスに変換する。
    /// </summary>
    public static string GetServerRelativeUrl(string url)
    {
        var uri = ParseUri(url);
        return Uri.UnescapeDataString(uri.AbsolutePath);
    }

    /// <summary>
    /// リソースURLからサイトURL（ルートまたは /sites|/teams）を導出する。
    /// </summary>
    public static string GetSiteUrl(string url)
    {
        var uri = ParseUri(url);
        var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var sitePath = string.Empty;

        if (segments.Length >= 2
            && (segments[0].Equals("sites", StringComparison.OrdinalIgnoreCase)
                || segments[0].Equals("teams", StringComparison.OrdinalIgnoreCase)))
        {
            sitePath = $"/{segments[0]}/{segments[1]}";
        }

        return $"{uri.Scheme}://{uri.Host}{sitePath}";
    }

    /// <summary>
    /// URLからファイル名部分を取り出す。
    /// </summary>
    public static string GetFileName(string url)
    {
        var uri = ParseUri(url);
        return Path.GetFileName(uri.LocalPath);
    }

    /// <summary>
    /// URLがフォルダを指すかどうかを簡易判定する。
    /// </summary>
    public static bool LooksLikeFolderUrl(string url)
    {
        var uri = ParseUri(url);
        var path = uri.AbsolutePath;
        return path.EndsWith("/", StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(Path.GetExtension(path));
    }

    /// <summary>
    /// ベースURLと相対セグメントを結合する。
    /// </summary>
    public static string CombineUrl(string baseUrl, string relative)
    {
        if (!baseUrl.EndsWith("/", StringComparison.Ordinal))
        {
            return $"{baseUrl}/{relative}";
        }

        return baseUrl + relative;
    }

    /// <summary>
    /// ベースURLと複数の相対セグメントを結合する。
    /// </summary>
    public static string CombineUrl(string baseUrl, string segment1, string segment2)
    {
        var combined = CombineUrl(baseUrl, segment1);
        return CombineUrl(combined, segment2);
    }

    private static Uri ParseUri(string value)
    {
        if (TryParseUri(value, out var uri))
        {
            return uri;
        }

        throw new InvalidOperationException($"Invalid URL: {value}");
    }

    private static bool TryParseUri(string value, out Uri uri)
    {
        if (Uri.TryCreate(value, UriKind.Absolute, out uri))
        {
            return true;
        }

        var escaped = Uri.EscapeUriString(value);
        return Uri.TryCreate(escaped, UriKind.Absolute, out uri);
    }
}
