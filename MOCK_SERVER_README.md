# Mock SPO Server - テストガイド

SharePoint Online REST APIをエミュレートするローカルモックサーバーです。  
実際のSPOに接続せずに、SpoCli の動作をテストできます。

## セットアップ

### 1. テストデータの作成

```powershell
.\create-test-data.ps1
```

これにより、以下の構造が `testdata/` フォルダに作成されます：

```
testdata/
  dirA/
    a.txt
    dirB/
      b.txt
      dirC/
        c.txt
```

### 2. モックサーバーの起動

```powershell
dotnet run --project MockSpoServer.csproj
```

サーバーは `http://localhost:5000` で起動します。

## SpoCli でのテスト

### フォルダ一覧表示

```powershell
.\bin\Debug\net10.0\SpoCli.exe ls "http://localhost:5000/sites/testsite/Shared Documents/dirA/"
```

### 再帰的ダウンロード

```powershell
# 出力ディレクトリを作成
New-Item -ItemType Directory -Path "output" -Force

# 再帰的にダウンロード
.\bin\Debug\net10.0\SpoCli.exe --recursive cp "http://localhost:5000/sites/testsite/Shared Documents/dirA/" ./output/
```

期待される結果：

```
output/
  a.txt
  dirB/
    b.txt
    dirC/
      c.txt
```

### ファイル単体ダウンロード

```powershell
.\bin\Debug\net10.0\SpoCli.exe cp "http://localhost:5000/sites/testsite/Shared Documents/dirA/a.txt" ./output/
```

## 仕組み

- `testdata/` フォルダがSharePointドキュメントライブラリとして機能
- REST API エンドポイントをエミュレート：
  - `/_api/web/GetFolderByServerRelativeUrl()` - フォルダ情報
  - `/_api/web/GetFileByServerRelativeUrl()/$value` - ファイルダウンロード
- 認証は不要（Bearer トークンのチェックなし）

## トラブルシューティング

### SpoCli が認証エラーを出す場合

SpoCli は通常 Azure AD 認証を要求しますが、モックサーバーは認証不要です。  
`SpoAuth.cs` を修正して、localhost の場合は認証をスキップするロジックを追加する必要があります。

### ポート 5000 が使用中の場合

`MockSpoServerProgram.cs` の最後の行を変更：

```csharp
app.Run("http://localhost:5001");  // ポート番号を変更
```
