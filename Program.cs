using System;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

/// <summary>
/// CLIのエントリポイントとコマンド振り分け。
/// </summary>
static class Program
{
	private const int ExitSuccess = 0;
	private const int ExitUsageError = 2;
	private const int ExitAuthError = 3;
	private const int ExitUnhandledError = 10;

	/// <summary>
	/// 引数を解析して各コマンドに振り分ける。
	/// </summary>
	public static async Task<int> Main(string[] args)
	{
		if (args.Length == 0 || IsHelp(args[0]))
		{
			PrintHelp();
			return ExitSuccess;
		}

		var command = args[0].ToLowerInvariant();
		var rest = args.Length > 1 ? args[1..] : Array.Empty<string>();

		try
		{
			return command switch
			{
				"login" => await HandleLoginAsync(rest),
				"ls" => await HandleListAsync(rest),
				"cp" => await HandleCopyAsync(rest),
				_ => HandleUnknown(command)
			};
		}
		catch (MsalUiRequiredException)
		{
			Console.Error.WriteLine("Authentication required. Run: spo-cli login --site <site-url>");
			return ExitAuthError;
		}
		catch (InvalidOperationException ex)
		{
			Console.Error.WriteLine(ex.Message);
			return ExitUsageError;
		}
		catch (Exception ex)
		{
			Console.Error.WriteLine($"Error: {ex.Message}");
			return ExitUnhandledError;
		}
	}

	/// <summary>
	/// 対話ログインを実行してトークンをキャッシュする。
	/// </summary>
	private static async Task<int> HandleLoginAsync(string[] args)
	{
		string? siteUrl = null;
		for (var i = 0; i < args.Length; i++)
		{
			if (args[i].Equals("--site", StringComparison.OrdinalIgnoreCase))
			{
				if (i + 1 >= args.Length)
				{
					throw new InvalidOperationException("Usage: spo-cli login --site <site-url>");
				}
				siteUrl = args[i + 1];
				i++;
			}
			else if (args[i].Equals("--mfa", StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}
			else
			{
				throw new InvalidOperationException($"Unknown option: {args[i]}");
			}
		}

		if (string.IsNullOrWhiteSpace(siteUrl))
		{
			siteUrl = Environment.GetEnvironmentVariable("SPO_SITE");
			if (string.IsNullOrWhiteSpace(siteUrl))
			{
				throw new InvalidOperationException("Usage: spo-cli login --site <site-url>");
			}
		}

		var config = SpoCliConfig.Load();
		var auth = await SpoAuth.CreateAsync(config);
		await auth.AcquireTokenResultAsync(siteUrl, interactive: true, prompt: Prompt.ForceLogin);

		var settings = SpoCliSettings.Load();
		settings.DefaultRoot = siteUrl;
		settings.Save();

		Console.WriteLine("Login succeeded.");
		Console.WriteLine($"Default root: {siteUrl}");
		return ExitSuccess;
	}

	/// <summary>
	/// SharePoint URL配下のフォルダとファイルを一覧表示する。
	/// </summary>
	private static async Task<int> HandleListAsync(string[] args)
	{
		if (args.Length != 1)
		{
			Console.Error.WriteLine("Usage: spo-cli ls <site-or-folder-url>");
			return ExitUsageError;
		}

		var url = UrlHelpers.ExpandSpoShorthand(args[0]);
		if (!UrlHelpers.IsSpoUrl(url))
		{
			Console.Error.WriteLine("ls requires a SharePoint Online URL.");
			return ExitUsageError;
		}

		var config = SpoCliConfig.Load();
		var auth = await SpoAuth.CreateAsync(config);
		var ops = new SpoOperations(auth);
		await ops.ListAsync(url);
		return ExitSuccess;
	}

	/// <summary>
	/// URL判定によりSPOとローカル間のコピーを実行する。
	/// </summary>
	private static async Task<int> HandleCopyAsync(string[] args)
	{
		if (args.Length != 2)
		{
			Console.Error.WriteLine("Usage: spo-cli cp <from> <to>");
			return ExitUsageError;
		}

		var from = UrlHelpers.ExpandSpoShorthand(args[0]);
		var to = UrlHelpers.ExpandSpoShorthand(args[1]);
		var fromIsSpo = UrlHelpers.IsSpoUrl(from);
		var toIsSpo = UrlHelpers.IsSpoUrl(to);

		if (!fromIsSpo && !toIsSpo)
		{
			Console.Error.WriteLine("cp requires at least one SharePoint Online URL.");
			return ExitUsageError;
		}

		var config = SpoCliConfig.Load();
		var auth = await SpoAuth.CreateAsync(config);
		var ops = new SpoOperations(auth);

		if (fromIsSpo && !toIsSpo)
		{
			await ops.DownloadAsync(from, to);
			return ExitSuccess;
		}

		if (!fromIsSpo && toIsSpo)
		{
			await ops.UploadAsync(from, to);
			return ExitSuccess;
		}

		await ops.CopyAsync(from, to);
		return ExitSuccess;
	}

	/// <summary>
	/// 不明なコマンド向けにヘルプを表示する。
	/// </summary>
	private static int HandleUnknown(string command)
	{
		Console.Error.WriteLine($"Unknown command: {command}");
		PrintHelp();
		return ExitUsageError;
	}

	/// <summary>
	/// ヘルプフラグかどうかを判定する。
	/// </summary>
	private static bool IsHelp(string arg)
	{
		return arg.Equals("-h", StringComparison.OrdinalIgnoreCase)
			|| arg.Equals("--help", StringComparison.OrdinalIgnoreCase)
			|| arg.Equals("help", StringComparison.OrdinalIgnoreCase);
	}

	/// <summary>
	/// CLIの使い方を表示する。
	/// </summary>
	private static void PrintHelp()
	{
		Console.WriteLine("spo-cli - SharePoint Online CLI (CSOM)");
		Console.WriteLine();
		Console.WriteLine("Commands:");
		Console.WriteLine("  login --site <site-url>        Login and cache token");
		Console.WriteLine("  ls <site-or-folder-url>        List files/folders");
		Console.WriteLine("  cp <from> <to>                 Copy file (SPO <-> local or SPO <-> SPO)");
		Console.WriteLine();
		Console.WriteLine("Env:");
		Console.WriteLine("  SPO_CLIENT_ID                  Entra ID app client ID");
		Console.WriteLine("  SPO_TENANT_ID                   Tenant ID (optional, default: organizations)");
		Console.WriteLine("  SPO_SITE                        Default site URL for login (optional)");
		Console.WriteLine();
		Console.WriteLine("Examples:");
		Console.WriteLine("  spo-cli login --site https://contoso.sharepoint.com");
		Console.WriteLine("  spo-cli ls https://contoso.sharepoint.com/sites/demo/Shared%20Documents");
		Console.WriteLine("  spo-cli cp https://contoso.sharepoint.com/sites/demo/Shared%20Documents/a.txt .\\a.txt");
	}
}
