using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

/// <summary>
/// 一覧表示とコピーのCSOM操作を実装する。
/// </summary>
internal sealed class SpoOperations
{
    private readonly SpoAuth _auth;

    public SpoOperations(SpoAuth auth)
    {
        _auth = auth;
    }

    /// <summary>
    /// 対象URL配下のフォルダとファイルを一覧表示する。
    /// </summary>
    public async Task ListAsync(string targetUrl)
    {
        var siteUrl = UrlHelpers.GetSiteUrl(targetUrl);
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false);

        var folderUrl = UrlHelpers.GetServerRelativeUrl(targetUrl);
        var folder = context.Web.GetFolderByServerRelativeUrl(folderUrl);
        context.Load(folder, f => f.Name, f => f.ServerRelativeUrl);
        context.Load(folder.Folders,
            folders => folders.Include(f => f.Name, f => f.ServerRelativeUrl, f => f.TimeLastModified));
        context.Load(folder.Files,
            files => files.Include(f => f.Name, f => f.ServerRelativeUrl, f => f.TimeLastModified, f => f.Length));
        context.ExecuteQuery();

        foreach (var subFolder in folder.Folders)
        {
            Console.WriteLine($"[D] {subFolder.Name}\t{subFolder.TimeLastModified:yyyy-MM-dd HH:mm}");
        }

        foreach (var file in folder.Files)
        {
            Console.WriteLine($"[F] {file.Name}\t{file.Length}\t{file.TimeLastModified:yyyy-MM-dd HH:mm}");
        }
    }

    /// <summary>
    /// SPOからローカルへファイルをダウンロードする。
    /// </summary>
    public async Task DownloadAsync(string fromUrl, string toPath)
    {
        var siteUrl = UrlHelpers.GetSiteUrl(fromUrl);
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false);

        var fileUrl = UrlHelpers.GetServerRelativeUrl(fromUrl);
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        var streamResult = file.OpenBinaryStream();
        context.Load(file);
        context.ExecuteQuery();
        using var sourceStream = streamResult.Value;

        var localPath = ResolveLocalPathForDownload(fromUrl, toPath);
        var directory = Path.GetDirectoryName(localPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var targetStream = System.IO.File.Create(localPath);
        sourceStream.CopyTo(targetStream);
    }

    /// <summary>
    /// ローカルのファイルをSPOにアップロードする。
    /// </summary>
    public async Task UploadAsync(string fromPath, string toUrl)
    {
        if (!System.IO.File.Exists(fromPath))
        {
            throw new FileNotFoundException("Source file not found.", fromPath);
        }

        var targetFileUrl = NormalizeTargetFileUrl(fromPath, toUrl);
        var siteUrl = UrlHelpers.GetSiteUrl(targetFileUrl);
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false);

        var serverRelative = UrlHelpers.GetServerRelativeUrl(targetFileUrl);
        var folderPath = serverRelative[..serverRelative.LastIndexOf('/')];
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            folderPath = "/";
        }

        var folder = context.Web.GetFolderByServerRelativeUrl(folderPath);
        using var stream = System.IO.File.OpenRead(fromPath);
        var info = new FileCreationInformation
        {
            ContentStream = stream,
            Url = Path.GetFileName(serverRelative),
            Overwrite = true
        };

        folder.Files.Add(info);
        context.ExecuteQuery();
    }

    /// <summary>
    /// SPO間でコピーし、テナントが異なる場合はダウンロード+アップロードに切り替える。
    /// </summary>
    public async Task CopyAsync(string fromUrl, string toUrl)
    {
        var sourceRoot = UrlHelpers.GetTenantRoot(fromUrl);
        var targetRoot = UrlHelpers.GetTenantRoot(toUrl);

        if (sourceRoot.Equals(targetRoot, StringComparison.OrdinalIgnoreCase))
        {
            await CopyWithinTenantAsync(fromUrl, toUrl);
            return;
        }

        var tempFile = Path.GetTempFileName();
        try
        {
            await DownloadAsync(fromUrl, tempFile);
            await UploadAsync(tempFile, toUrl);
        }
        finally
        {
            System.IO.File.Delete(tempFile);
        }
    }

    /// <summary>
    /// 同一テナント内のファイルコピーにMoveCopyUtilを使う。
    /// </summary>
    private async Task CopyWithinTenantAsync(string fromUrl, string toUrl)
    {
        var targetFileUrl = NormalizeTargetFileUrl(fromUrl, toUrl);
        var siteUrl = UrlHelpers.GetSiteUrl(fromUrl);
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false);

        var sourcePath = ResourcePath.FromDecodedUrl(UrlHelpers.GetServerRelativeUrl(fromUrl));
        var targetPath = ResourcePath.FromDecodedUrl(UrlHelpers.GetServerRelativeUrl(targetFileUrl));
        var options = new MoveCopyOptions
        {
            KeepBoth = false,
            ResetAuthorAndCreatedOnCopy = false
        };

        MoveCopyUtil.CopyFileByPath(context, sourcePath, targetPath, overwrite: true, options: options);
        context.ExecuteQuery();
    }

    /// <summary>
    /// 保存先がディレクトリか判定し、必要ならファイル名を補う。
    /// </summary>
    private static string ResolveLocalPathForDownload(string fromUrl, string toPath)
    {
        if (Directory.Exists(toPath)
            || toPath.EndsWith(Path.DirectorySeparatorChar)
            || toPath.EndsWith(Path.AltDirectorySeparatorChar))
        {
            return Path.Combine(toPath, UrlHelpers.GetFileName(fromUrl));
        }

        return toPath;
    }

    /// <summary>
    /// フォルダURLが指定された場合は、元のファイル名を付与する。
    /// </summary>
    private static string NormalizeTargetFileUrl(string fromPath, string toUrl)
    {
        if (!UrlHelpers.LooksLikeFolderUrl(toUrl))
        {
            return toUrl;
        }

        var fileName = UrlHelpers.IsSpoUrl(fromPath)
            ? UrlHelpers.GetFileName(fromPath)
            : Path.GetFileName(fromPath);
        return UrlHelpers.CombineUrl(toUrl, fileName);
    }
}
