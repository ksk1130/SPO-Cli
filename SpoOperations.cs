using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;

/// <summary>
/// ダウンロード対象アイテムの情報
/// </summary>
internal record DownloadItem(string ServerRelativeUrl, string LocalPath, long FileLength, string FileName);

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
    /// ダウンロードフロー：①リスト作成（CSOM） -> ②一括ダウンロード（CSOM）
    /// すべて CSOM 経由で処理することで、'#' 等の特殊文字を含むパスも正しく動作する。
    /// </summary>
    public async Task DownloadAsync(string fromUrl, string toPath, bool recursive = false)
    {
        var siteUrl = UrlHelpers.GetSiteUrl(fromUrl);
        var serverRelativeUrl = UrlHelpers.GetServerRelativeUrl(fromUrl);
        var fileName = UrlHelpers.GetFileName(fromUrl);

        // ファイル拡張子があればフォルダ試行を飛ばして直接ファイルとして処理
        bool looksLikeFile = !string.IsNullOrEmpty(Path.GetExtension(fileName));

        var downloadList = new List<DownloadItem>();
        bool isFolder = false;

        // CSOM コンテキストをフォルダ列挙・ファイルダウンロード両方で共用する
        using var csomContext = await _auth.CreateContextAsync(siteUrl, interactive: false, showExpiresOn: false);

        if (!looksLikeFile)
        {
            try
            {
                BuildDownloadListWithCsom(csomContext, serverRelativeUrl, serverRelativeUrl, toPath, downloadList, recursive ? 3 : 0);
                isFolder = true;
            }
            catch
            {
                isFolder = false;
            }
        }

        // ② フォルダでない場合はファイルとして処理（CSOM 経由）
        if (!isFolder)
        {
            try
            {
                var resolvedPath = ResolveLocalPathForDownload(fromUrl, toPath);
                var resolvedDir = Path.GetDirectoryName(resolvedPath);
                if (!string.IsNullOrEmpty(resolvedDir))
                {
                    Directory.CreateDirectory(resolvedDir);
                }

                // CSOM でファイルサイズを取得してダウンロード
                var fileRef = csomContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
                csomContext.Load(fileRef, f => f.Length);
                csomContext.ExecuteQuery();
                var fileLength = fileRef.Length;

                DownloadFileWithCsom(csomContext, serverRelativeUrl, resolvedPath, fileLength, fileName);
                Console.WriteLine("Download completed successfully.");
                return;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Error] Neither file nor folder found: {serverRelativeUrl}");
                Console.Error.WriteLine($"[Error] {ex.Message}");
                throw;
            }
        }

        // ③ フォルダからのダウンロード（CSOM 経由）
        if (downloadList.Count == 0)
        {
            Console.WriteLine("No files found to download.");
            return;
        }

        PrintDownloadList(downloadList);

        if (recursive && !ConfirmRecursiveDownload(downloadList.Count))
        {
            Console.WriteLine("Download canceled.");
            return;
        }

        foreach (var item in downloadList)
        {
            var dir = Path.GetDirectoryName(item.LocalPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            Console.WriteLine(item.LocalPath);
            DownloadFileWithCsom(csomContext, item.ServerRelativeUrl, item.LocalPath, item.FileLength, item.FileName);
        }
        Console.WriteLine("Download completed successfully.");
    }

    /// <summary>
    /// CSOM を使ってファイルをダウンロードする。
    /// GetFileByServerRelativePath(ResourcePath.FromDecodedUrl) を使用するため、
    /// '#' 等の特殊文字を含むパスもリクエストボディ経由で正しく処理できる。
    /// </summary>
    private static void DownloadFileWithCsom(
        Microsoft.SharePoint.Client.ClientContext context,
        string serverRelativeUrl,
        string localPath,
        long fileLength,
        string fileName)
    {
        try
        {
            var fileRef = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            var streamResult = fileRef.OpenBinaryStream();
            context.ExecuteQuery();

            using var sourceStream = streamResult.Value;
            using var targetStream = System.IO.File.Create(localPath);
            CopyStreamWithProgress(sourceStream, targetStream, fileLength, fileName);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[Error] Failed to download: {serverRelativeUrl}");
            Console.Error.WriteLine($"[Error] Exception: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// CSOM を使ってフォルダ内のファイル・サブフォルダを列挙し、ダウンロードリストを構築する。
    /// CSOM はパスを URL ではなくリクエストボディに埋め込むため、'#' 等の特殊文字を含む
    /// フォルダ名でも正しく動作する。
    /// </summary>
    private static void BuildDownloadListWithCsom(
        Microsoft.SharePoint.Client.ClientContext context,
        string serverRelativeUrl,
        string baseServerRelativeUrl,
        string baseLocalDir,
        List<DownloadItem> list,
        int maxDepth,
        int currentDepth = 0)
    {
        serverRelativeUrl = serverRelativeUrl.TrimEnd('/');
        baseServerRelativeUrl = baseServerRelativeUrl.TrimEnd('/');

        var relativeFolderPath = GetRelativePath(baseServerRelativeUrl, serverRelativeUrl);
        var localDir = string.IsNullOrEmpty(relativeFolderPath)
            ? baseLocalDir
            : Path.Combine(baseLocalDir, relativeFolderPath);

        // ResourcePath.FromDecodedUrl を使うことで、CSOM がパスを XML ボディの文字列として送信する。
        // '#' は XML 文字列内で特殊文字でないため、フラグメント誤認識の問題が発生しない。
        var folder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
        context.Load(folder.Files,
            files => files.Include(f => f.Name, f => f.ServerRelativeUrl, f => f.Length));
        if (currentDepth < maxDepth)
        {
            context.Load(folder.Folders,
                folders => folders.Include(f => f.Name, f => f.ServerRelativeUrl));
        }
        context.ExecuteQuery();

        // ファイル情報を先に収集（ExecuteQuery 後にプロキシ値を読み取る）
        var fileInfos = folder.Files.AsEnumerable()
            .Select(f => (name: f.Name, url: f.ServerRelativeUrl, length: f.Length))
            .ToList();

        foreach (var (name, url, length) in fileInfos)
        {
            var localPath = Path.Combine(localDir, name);
            list.Add(new DownloadItem(url, localPath, length, name));
        }

        if (currentDepth < maxDepth)
        {
            // サブフォルダ情報も収集してから再帰（プロキシオブジェクトを次の ExecuteQuery 前に退避）
            var subFolders = folder.Folders.AsEnumerable()
                .Select(f => (name: f.Name, url: f.ServerRelativeUrl))
                .Where(f => f.name != "Forms")
                .ToList();

            foreach (var (_, url) in subFolders)
            {
                BuildDownloadListWithCsom(context, url, baseServerRelativeUrl, baseLocalDir, list, maxDepth, currentDepth + 1);
            }
        }
    }
    /// <summary>
    /// PascalCase / camelCase の両キーを試みて string を返す。
    /// </summary>
    private static string? GetStringProperty(System.Text.Json.JsonElement el, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (el.TryGetProperty(key, out var prop) && prop.ValueKind == System.Text.Json.JsonValueKind.String)
                return prop.GetString();
        }
        return null;
    }

    /// <summary>
    /// PascalCase / camelCase の両キーを試みて long を返す。
    /// </summary>
    private static long GetInt64Property(System.Text.Json.JsonElement el, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (el.TryGetProperty(key, out var prop))
            {
                if (prop.ValueKind == System.Text.Json.JsonValueKind.Number)
                    return prop.GetInt64();
                // SPO は Length を文字列で返すこともある
                if (prop.ValueKind == System.Text.Json.JsonValueKind.String
                    && long.TryParse(prop.GetString(), out var parsed))
                    return parsed;
            }
        }
        return 0;
    }

    /// <summary>
    /// ダウンロード予定のファイル一覧を表示する。
    /// </summary>
    private static void PrintDownloadList(IReadOnlyList<DownloadItem> downloadList)
    {
        Console.WriteLine($"\nFiles to download ({downloadList.Count}):");
        for (int i = 0; i < downloadList.Count; i++)
        {
            var item = downloadList[i];
            Console.WriteLine($"{i + 1,2}. {item.ServerRelativeUrl} -> {item.LocalPath}");
        }
        Console.WriteLine();
    }

    /// <summary>
    /// 再帰ダウンロード実行前の確認を行う。
    /// </summary>
    private static bool ConfirmRecursiveDownload(int fileCount)
    {
        Console.Write($"Download {fileCount} files recursively? (y/N): ");
        var input = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        return input.Trim().Equals("y", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// ベースパスからの相対パスを計算する。
    /// </summary>
    private static string GetRelativePath(string basePath, string fullPath)
    {
        basePath = basePath.TrimEnd('/');
        fullPath = fullPath.TrimEnd('/');

        if (fullPath.Equals(basePath, StringComparison.OrdinalIgnoreCase))
        {
            return string.Empty;
        }

        if (fullPath.StartsWith(basePath + "/", StringComparison.OrdinalIgnoreCase))
        {
            var relative = fullPath.Substring(basePath.Length + 1);
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        throw new InvalidOperationException($"Path '{fullPath}' is not under base path '{basePath}'");
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
    /// URL が localhost を指しているかチェックする。
    /// </summary>
    private static bool IsLocalhost(string url)
    {
        if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                || uri.Host.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    /// <summary>
    /// localhost のテストデータからダウンロード（モックサーバー用）
    /// </summary>
    private void DownloadFromLocalhost(string fromUrl, string toPath, bool recursive = false)
    {
        var uri = new Uri(fromUrl);
        var localPath = MapLocalhostUrlToLocalPath(uri.AbsolutePath);

        Console.Error.WriteLine($"[Debug] Looking for testdata at: {Path.GetFullPath(localPath)}");

        if (System.IO.File.Exists(localPath))
        {
            // ファイルの場合
            var destination = ResolveLocalPathForDownload(fromUrl, toPath);
            var destinationDir = Path.GetDirectoryName(destination);
            if (!string.IsNullOrWhiteSpace(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }
            System.IO.File.Copy(localPath, destination, overwrite: true);
            Console.Error.WriteLine($"Downloaded: {Path.GetFileName(destination)}");
        }
        else if (Directory.Exists(localPath))
        {
            // フォルダの場合
            if (recursive)
            {
                DownloadFolderLocalhostRecursive(localPath, toPath);
            }
            else
            {
                DownloadFolderLocalhostFiles(localPath, toPath);
            }
        }
        else
        {
            throw new InvalidOperationException($"Path not found: {Path.GetFullPath(localPath)}");
        }
    }

    /// <summary>
    /// localhostのフォルダ直下のファイルをダウンロード
    /// </summary>
    private void DownloadFolderLocalhostFiles(string sourceFolder, string targetPath)
    {
        Directory.CreateDirectory(targetPath);

        var files = Directory.GetFiles(sourceFolder);
        foreach (var file in files)
        {
            var fileName = Path.GetFileName(file);
            var destination = Path.Combine(targetPath, fileName);
            System.IO.File.Copy(file, destination, overwrite: true);
            Console.Error.WriteLine($"Downloaded: {fileName}");
        }
    }

    /// <summary>
    /// localhostのフォルダを再帰的にダウンロード
    /// </summary>
    private void DownloadFolderLocalhostRecursive(string sourceFolder, string targetPath, int currentDepth = 0, int maxDepth = 3)
    {
        Directory.CreateDirectory(targetPath);

        if (currentDepth > maxDepth)
        {
            return;
        }

        // 直下のすべてのファイルをダウンロード
        var files = Directory.GetFiles(sourceFolder);
        foreach (var file in files)
        {
            var fileName = Path.GetFileName(file);
            var destination = Path.Combine(targetPath, fileName);
            System.IO.File.Copy(file, destination, overwrite: true);
            Console.Error.WriteLine($"Downloaded: {fileName}");
        }

        // サブフォルダを再帰処理
        if (currentDepth < maxDepth)
        {
            var subFolders = Directory.GetDirectories(sourceFolder);
            foreach (var subFolder in subFolders)
            {
                var folderName = Path.GetFileName(subFolder);
                var subTargetPath = Path.Combine(targetPath, folderName);
                DownloadFolderLocalhostRecursive(subFolder, subTargetPath, currentDepth + 1, maxDepth);
            }
        }
    }

    /// <summary>
    /// localhost URL のパスをローカルファイルシステムパスにマッピング
    /// </summary>
    private static string MapLocalhostUrlToLocalPath(string serverRelativePath)
    {
        // /sites/testsite/Shared Documents/dirA -> dirA
        var normalized = serverRelativePath
            .Replace("/sites/testsite/Shared Documents/", "")
            .Replace("/sites/testsite/Shared%20Documents/", "")
            .TrimStart('/')
            .Replace("/", Path.DirectorySeparatorChar.ToString())
            .Replace("%20", " ");

        // 実行可能ファイルのディレクトリを基準に testdata を探す
        var exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        var exeDir = Path.GetDirectoryName(exePath) ?? ".";

        // publish フォルダから実行される場合: bin\publish\ -> ..\..\testdata\ で SPO-Cli\testdata に到達
        var testDataPath = Path.Combine(exeDir, "..", "..", "testdata", normalized);

        return Path.GetFullPath(testDataPath);
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
