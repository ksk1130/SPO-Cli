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
		//Console.Error.WriteLine($"[Debug] SpoCliSettings.Load() -> DefaultRoot={settings.DefaultRoot}");
		
		if (string.IsNullOrWhiteSpace(settings.DefaultRoot))
		{
			throw new InvalidOperationException("spo:// を使うには先に login でサイトを指定してください。");
		}

		var relative = value.Substring(6);
		//Console.Error.WriteLine($"[Debug] ExpandSpoShorthand: input={value}, DefaultRoot={settings.DefaultRoot}, relative={relative}");
		
		string result;
		if (string.IsNullOrWhiteSpace(relative))
		{
			result = CombineUrl(settings.DefaultRoot, "Shared Documents");
		}
		else
		{
			result = CombineUrl(settings.DefaultRoot, "Shared Documents", relative);
		}
		
		//Console.Error.WriteLine($"[Debug] ExpandSpoShorthand: result={result}");
		return result;
	}

	/// <summary>
	/// SharePoint Online URLかどうかを判定する。
	/// localhostも対応（モックサーバー用）
	/// </summary>
	public static bool IsSpoUrl(string value)
	{
        Uri uri;
        return TryParseUri(value, out uri)
			&& (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp)
            && (uri.Host.IndexOf("sharepoint.com", StringComparison.OrdinalIgnoreCase) >= 0
				|| uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
				|| uri.Host.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase));
	}

    /// <summary>
    /// サイトURLのスキームとホストを返す。
    /// </summary>
    public static string GetTenantRoot(string url)
    {
        var uri = ParseUri(url);
        return uri.Scheme + "://" + uri.Authority;
    }

    /// <summary>
    /// 絶対URLをサーバー相対パスに変換する。
    /// </summary>
    public static string GetServerRelativeUrl(string url)
    {
        var uri = ParseUri(url);
        var result = Uri.UnescapeDataString(uri.AbsolutePath);
        //Console.Error.WriteLine($"[Debug] GetServerRelativeUrl: url={url} -> AbsolutePath={uri.AbsolutePath} -> result={result}");
        return result;
    }

    /// <summary>
    /// リソースURLからサイトURL（ルートまたは /sites|/teams）を導出する。
    /// </summary>
    public static string GetSiteUrl(string url)
    {
        var uri = ParseUri(url);
        var segments = uri.AbsolutePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var sitePath = string.Empty;

        if (segments.Length >= 2
            && (segments[0].Equals("sites", StringComparison.OrdinalIgnoreCase)
                || segments[0].Equals("teams", StringComparison.OrdinalIgnoreCase)))
        {
            sitePath = "/" + segments[0] + "/" + segments[1];
        }

        var result = uri.Scheme + "://" + uri.Authority + sitePath;
        //Console.Error.WriteLine($"[Debug] GetSiteUrl: url={url} -> segments={string.Join(",", segments)} -> result={result}");
        return result;
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
            return baseUrl + "/" + relative;
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
        Uri uri;
        if (TryParseUri(value, out uri))
        {
            return uri;
        }

        throw new InvalidOperationException("Invalid URL: " + value);
    }

    private static bool TryParseUri(string value, out Uri uri)
    {
        // '#' は URI のフラグメント区切り文字。パス内のフォルダ／ファイル名に '#' が含まれる場合、
        // Uri.TryCreate に渡す前に %23 に置換しないと '#' 以降がパスから切り捨てられる。
        // SPO の URL にフラグメントが使われることはないため、全ての '#' を安全に %23 へ変換する。
        var sanitized = value.Replace("#", "%23");
        if (Uri.TryCreate(sanitized, UriKind.Absolute, out uri))
        {
            return true;
        }

        var escaped = Uri.EscapeUriString(sanitized);
        return Uri.TryCreate(escaped, UriKind.Absolute, out uri);
    }
}
