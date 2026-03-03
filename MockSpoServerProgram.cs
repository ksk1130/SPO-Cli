using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;

/// <summary>
/// SharePoint Online REST APIをエミュレートするモックサーバー
/// ローカルの testdata/ フォルダをSPOのドキュメントライブラリとして扱う
/// </summary>
var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

// テストデータのルートディレクトリ
var testDataRoot = Path.Combine(Directory.GetCurrentDirectory(), "testdata");
Directory.CreateDirectory(testDataRoot);

Console.WriteLine($"Mock SPO Server starting...");
Console.WriteLine($"Test data root: {testDataRoot}");
Console.WriteLine($"Listening on: http://localhost:5000");
Console.WriteLine($"Mock site URL: http://localhost:5000/sites/testsite");
Console.WriteLine();
Console.WriteLine("Example usage:");
Console.WriteLine("  spocli --recursive cp http://localhost:5000/sites/testsite/Shared%20Documents/dirA/ .");
Console.WriteLine();

// CSOM contextinfo エンドポイント（認証用）
app.MapPost("/sites/{site}/_api/contextinfo", () =>
{
    return Results.Json(new
    {
        d = new
        {
            GetContextWebInformation = new
            {
                FormDigestValue = "mock-digest-value",
                FormDigestTimeoutSeconds = 1800
            }
        }
    });
});

// フォルダ情報取得エンドポイント
app.MapGet("/sites/{site}/{*apiPath}", 
    (string site, string apiPath) =>
{
    // GetFolderByServerRelativeUrl('path')/Files?$select=Name,Length
    var filesMatch = System.Text.RegularExpressions.Regex.Match(
        apiPath,
        @"_api/web/GetFolderByServerRelativeUrl\('([^']+)'\)/Files");

    if (filesMatch.Success)
    {
        var serverRelativePath = filesMatch.Groups[1].Value;
        var localPath = MapServerRelativeToLocal(serverRelativePath, testDataRoot);

        if (!Directory.Exists(localPath))
            return Results.NotFound(new { error = new { message = "Folder not found" } });

        var folderInfo = new DirectoryInfo(localPath);
        var files = folderInfo.GetFiles()
            .Select(f => new
            {
                Name = f.Name,
                Length = f.Length
            }).ToArray();

        return Results.Json(new { value = files });
    }

    // GetFolderByServerRelativeUrl('path')/Folders?$select=Name
    var foldersMatch = System.Text.RegularExpressions.Regex.Match(
        apiPath,
        @"_api/web/GetFolderByServerRelativeUrl\('([^']+)'\)/Folders");

    if (foldersMatch.Success)
    {
        var serverRelativePath = foldersMatch.Groups[1].Value;
        var localPath = MapServerRelativeToLocal(serverRelativePath, testDataRoot);

        if (!Directory.Exists(localPath))
            return Results.NotFound(new { error = new { message = "Folder not found" } });

        var folderInfo = new DirectoryInfo(localPath);
        var folders = folderInfo.GetDirectories()
            .Select(d => new { Name = d.Name }).ToArray();

        return Results.Json(new { value = folders });
    }

    // GetFolderByServerRelativeUrl('/Shared Documents/dirA') をパース（CSOM互換）
    var match = System.Text.RegularExpressions.Regex.Match(
        apiPath,
        @"_api/web/GetFolderByServerRelativeUrl\('([^']+)'\)$");
    
    if (match.Success)
    {
        var serverRelativePath = match.Groups[1].Value;
        var localPath = MapServerRelativeToLocal(serverRelativePath, testDataRoot);

        if (!Directory.Exists(localPath))
        {
            return Results.NotFound(new { error = new { message = $"Folder not found" } });
        }

        var folderInfo = new DirectoryInfo(localPath);
        var folders = folderInfo.GetDirectories()
            .Select(d => new
            {
                Name = d.Name,
                ServerRelativeUrl = MapLocalToServerRelative(d.FullName, testDataRoot, site),
                TimeLastModified = d.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }).ToArray();

        var files = folderInfo.GetFiles()
            .Select(f => new
            {
                Name = f.Name,
                ServerRelativeUrl = MapLocalToServerRelative(f.FullName, testDataRoot, site),
                Length = f.Length,
                TimeLastModified = f.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }).ToArray();

        return Results.Json(new
        {
            d = new
            {
                Name = folderInfo.Name,
                ServerRelativeUrl = MapLocalToServerRelative(localPath, testDataRoot, site),
                Folders = new { results = folders },
                Files = new { results = files }
            }
        });
    }

    // GetFileByServerRelativeUrl('/Shared Documents/dirA/a.txt')/$value をパース
    var fileValueMatch = System.Text.RegularExpressions.Regex.Match(
        apiPath,
        @"_api/web/GetFileByServerRelativeUrl\('([^']+)'\)/\$value");
    
    if (fileValueMatch.Success)
    {
        var serverRelativePath = fileValueMatch.Groups[1].Value;
        var localPath = MapServerRelativeToLocal(serverRelativePath, testDataRoot);

        if (!File.Exists(localPath))
            return Results.NotFound();

        var fileStream = File.OpenRead(localPath);
        return Results.Stream(fileStream, "application/octet-stream", Path.GetFileName(localPath));
    }

    // GetFileByServerRelativeUrl('/Shared Documents/dirA/a.txt') をパース
    var fileMetaMatch = System.Text.RegularExpressions.Regex.Match(
        apiPath,
        @"_api/web/GetFileByServerRelativeUrl\('([^']+)'\)$");
    
    if (fileMetaMatch.Success)
    {
        var serverRelativePath = fileMetaMatch.Groups[1].Value;
        var localPath = MapServerRelativeToLocal(serverRelativePath, testDataRoot);

        if (!File.Exists(localPath))
            return Results.NotFound();

        var fileInfo = new FileInfo(localPath);
        return Results.Json(new
        {
            d = new
            {
                Name = fileInfo.Name,
                Length = fileInfo.Length,
                ServerRelativeUrl = MapLocalToServerRelative(localPath, testDataRoot, site),
                TimeLastModified = fileInfo.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
        });
    }

    return Results.BadRequest(new { error = "Unknown API endpoint" });
});

// contextinfo エンドポイント（POST）
app.MapPost("/sites/{site}/{*apiPath}", (string site, string apiPath) =>
{
    if (apiPath.Contains("contextinfo"))
    {
        return Results.Json(new
        {
            d = new
            {
                GetContextWebInformation = new
                {
                    FormDigestValue = "mock-digest-value",
                    FormDigestTimeoutSeconds = 1800
                }
            }
        });
    }

    return Results.BadRequest(new { error = "Unknown API endpoint" });
});

app.Run("http://localhost:5000");

// サーバー相対パスをローカルパスにマッピング
static string MapServerRelativeToLocal(string serverRelativePath, string root)
{
    // /sites/{any}/Shared Documents/dirA -> testdata/dirA
    var normalized = serverRelativePath
        .Replace("/sites/testsite/Shared Documents/", "")      // 互換性用
        .Replace("/sites/testsite/Shared%20Documents/", "");    // 互換性用
    
    // まだ /sites/ が残ってる場合は任意のサイト名パターンなので正規表現で処理
    if (normalized.StartsWith("/sites/"))
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            normalized, 
            @"^/sites/[^/]+/Shared%20Documents/(.*)$");
        if (match.Success)
        {
            normalized = match.Groups[1].Value;
        }
        else
        {
            match = System.Text.RegularExpressions.Regex.Match(
                normalized,
                @"^/sites/[^/]+/Shared Documents/(.*)$");
            if (match.Success)
            {
                normalized = match.Groups[1].Value;
            }
        }
    }
    
    normalized = normalized
        .Replace("/", Path.DirectorySeparatorChar.ToString())
        .TrimStart(Path.DirectorySeparatorChar);

    return Path.Combine(root, normalized);
}

// ローカルパスをサーバー相対パスにマッピング
static string MapLocalToServerRelative(string localPath, string root, string site = "testsite")
{
    var relativePath = Path.GetRelativePath(root, localPath)
        .Replace(Path.DirectorySeparatorChar, '/');
    return $"/sites/{site}/Shared Documents/{relativePath}";
}
