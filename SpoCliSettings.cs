using System;
using System.IO;
using System.Text.Json;

/// <summary>
/// ログイン設定（デフォルトルートなど）を永続化する。
/// </summary>
internal sealed class SpoCliSettings
{
	private const string SettingsFileName = "config.json";

	public string? DefaultRoot { get; set; }

	private static string GetSettingsPath()
	{
		var dir = Path.Combine(
			Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
			"spo-cli");
		Directory.CreateDirectory(dir);
		return Path.Combine(dir, SettingsFileName);
	}

	/// <summary>
	/// 設定ファイルから読み込む。存在しない場合は空の設定を返す。
	/// </summary>
	public static SpoCliSettings Load()
	{
		var path = GetSettingsPath();
		if (!File.Exists(path))
		{
			return new SpoCliSettings();
		}

		try
		{
			var json = File.ReadAllText(path);
			return JsonSerializer.Deserialize<SpoCliSettings>(json) ?? new SpoCliSettings();
		}
		catch
		{
			return new SpoCliSettings();
		}
	}

	/// <summary>
	/// 設定ファイルに保存する。
	/// </summary>
	public void Save()
	{
		var path = GetSettingsPath();
		var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
		File.WriteAllText(path, json);
	}
}
