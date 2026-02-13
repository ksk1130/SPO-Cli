using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
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
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false, showExpiresOn: true);

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
        if (fromUrl.EndsWith("/", StringComparison.Ordinal))
        {
            await DownloadFolderFilesAsync(fromUrl, toPath);
            return;
        }

        var siteUrl = UrlHelpers.GetSiteUrl(fromUrl);
        var tokenResult = await _auth.AcquireTokenResultAsync(siteUrl, interactive: false);
        using var context = await _auth.CreateContextAsync(
            siteUrl,
            interactive: false,
            showExpiresOn: true,
            tokenResult: tokenResult);

        var fileUrl = UrlHelpers.GetServerRelativeUrl(fromUrl);
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        context.Load(file, f => f.Length, f => f.Name);
        context.ExecuteQuery();

        var localPath = ResolveLocalPathForDownload(fromUrl, toPath);
        var directory = Path.GetDirectoryName(localPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        await DownloadFileByServerRelativeUrlAsync(
            siteUrl,
            fileUrl,
            localPath,
            file.Length,
            file.Name,
            tokenResult);
    }

    /// <summary>
    /// フォルダ直下のファイルを再帰せずにダウンロードする。
    /// </summary>
    private async Task DownloadFolderFilesAsync(string folderUrl, string toPath)
    {
        if (!Directory.Exists(toPath)
            && !toPath.EndsWith(Path.DirectorySeparatorChar)
            && !toPath.EndsWith(Path.AltDirectorySeparatorChar))
        {
            throw new InvalidOperationException("Folder download requires a local directory path.");
        }

        Directory.CreateDirectory(toPath);

        var siteUrl = UrlHelpers.GetSiteUrl(folderUrl);
        var tokenResult = await _auth.AcquireTokenResultAsync(siteUrl, interactive: false);
        using var context = await _auth.CreateContextAsync(
            siteUrl,
            interactive: false,
            showExpiresOn: true,
            tokenResult: tokenResult);

        var serverRelativeFolder = UrlHelpers.GetServerRelativeUrl(folderUrl);
        var folder = context.Web.GetFolderByServerRelativeUrl(serverRelativeFolder);
        context.Load(folder, f => f.Name);
        context.Load(folder.Files, files => files.Include(f => f.Name, f => f.Length, f => f.ServerRelativeUrl));
        context.ExecuteQuery();

        foreach (var file in folder.Files)
        {
            var localPath = Path.Combine(toPath, file.Name);
            await DownloadFileByServerRelativeUrlAsync(
                siteUrl,
                file.ServerRelativeUrl,
                localPath,
                file.Length,
                file.Name,
                tokenResult);
        }
    }

    /// <summary>
    /// サーバー相対URLのファイルをRESTでストリーミングダウンロードする。
    /// </summary>
    private async Task DownloadFileByServerRelativeUrlAsync(
        string siteUrl,
        string serverRelativeUrl,
        string localPath,
        long fileLength,
        string fileName,
        AuthenticationResult tokenResult)
    {
        var encodedPath = Uri.EscapeDataString(serverRelativeUrl);
        var downloadUrl = $"{siteUrl}/_api/web/GetFileByServerRelativeUrl('{encodedPath}')/$value";
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", tokenResult.AccessToken);
        using var response = await client.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();
        using var sourceStream = await response.Content.ReadAsStreamAsync();

        using var targetStream = System.IO.File.Create(localPath);
        var totalBytes = response.Content.Headers.ContentLength ?? fileLength;
        CopyStreamWithProgress(sourceStream, targetStream, totalBytes, fileName);
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
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false, showExpiresOn: true);

        var serverRelative = UrlHelpers.GetServerRelativeUrl(targetFileUrl);
        var folderPath = serverRelative[..serverRelative.LastIndexOf('/')];
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            folderPath = "/";
        }

        var folder = context.Web.GetFolderByServerRelativeUrl(folderPath);
        var fileInfo = new FileInfo(fromPath);
        using var sourceStream = System.IO.File.OpenRead(fromPath);
        using var progressStream = new ProgressStream(sourceStream, fileInfo.Length, Path.GetFileName(fromPath));
        var info = new FileCreationInformation
        {
            ContentStream = progressStream,
            Url = Path.GetFileName(serverRelative),
            Overwrite = true
        };

        folder.Files.Add(info);
        context.ExecuteQuery();
        Console.Error.WriteLine();
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
        using var context = await _auth.CreateContextAsync(siteUrl, interactive: false, showExpiresOn: true);

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

    /// <summary>
    /// ストリームをコピーしながら進捗表示する。
    /// </summary>
    private static void CopyStreamWithProgress(Stream source, Stream destination, long totalBytes, string fileName)
    {
        const int bufferSize = 81920;
        var buffer = new byte[bufferSize];
        long totalRead = 0;
        int bytesRead;
        var lastPercent = -1;

        while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            destination.Write(buffer, 0, bytesRead);
            totalRead += bytesRead;

            var percent = (int)((double)totalRead / totalBytes * 100);
            if (percent != lastPercent)
            {
                DisplayProgress(fileName, totalRead, totalBytes, percent);
                lastPercent = percent;
            }
        }

        Console.Error.WriteLine();
    }

    /// <summary>
    /// 進捗バーを表示する。
    /// </summary>
    private static void DisplayProgress(string fileName, long current, long total, int percent)
    {
        const int barWidth = 30;
        var filled = (int)(barWidth * percent / 100);
        var bar = new string('#', filled) + new string('-', barWidth - filled);
        var currentMB = current / 1024.0 / 1024.0;
        var totalMB = total / 1024.0 / 1024.0;
        Console.Error.Write($"\r{fileName}: [{bar}] {percent}% {currentMB:F2}MB/{totalMB:F2}MB");
        Console.Error.Flush();
    }
}

/// <summary>
/// 進捗表示付きストリームラッパー（アップロード用）。
/// </summary>
internal sealed class ProgressStream : Stream
{
    private readonly Stream _baseStream;
    private readonly long _totalBytes;
    private readonly string _fileName;
    private long _bytesRead;
    private int _lastPercent = -1;

    public ProgressStream(Stream baseStream, long totalBytes, string fileName)
    {
        _baseStream = baseStream;
        _totalBytes = totalBytes;
        _fileName = fileName;
    }

    public override bool CanRead => _baseStream.CanRead;
    public override bool CanSeek => _baseStream.CanSeek;
    public override bool CanWrite => false;
    public override long Length => _baseStream.Length;
    public override long Position
    {
        get => _baseStream.Position;
        set => _baseStream.Position = value;
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        var bytesRead = _baseStream.Read(buffer, offset, count);
        _bytesRead += bytesRead;

        var percent = _totalBytes > 0 ? (int)((double)_bytesRead / _totalBytes * 100) : 0;
        if (percent != _lastPercent)
        {
            DisplayProgress(percent);
            _lastPercent = percent;
        }

        return bytesRead;
    }

    private void DisplayProgress(int percent)
    {
        const int barWidth = 30;
        var filled = (int)(barWidth * percent / 100);
        var bar = new string('#', filled) + new string('-', barWidth - filled);
        var currentMB = _bytesRead / 1024.0 / 1024.0;
        var totalMB = _totalBytes / 1024.0 / 1024.0;
        Console.Error.Write($"\r{_fileName}: [{bar}] {percent}% {currentMB:F2}MB/{totalMB:F2}MB");
        Console.Error.Flush();
    }

    public override void Flush() => _baseStream.Flush();
    public override long Seek(long offset, SeekOrigin origin) => _baseStream.Seek(offset, origin);
    public override void SetLength(long value) => _baseStream.SetLength(value);
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _baseStream.Dispose();
        }
        base.Dispose(disposing);
    }
}
