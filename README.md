# spo-cli

SharePoint Online のファイル操作をCLIで(aws s3コマンド風に)実行できるツールです。

## コマンド

- `login [--mfa]`
- `ls <site-or-folder-url>`
- `cp [--recursive] <from> <to>`

CSOMの対話ログインとMSALトークンキャッシュを使います。

日本語を含むURLは、生のURLとエンコード済みの両方に対応しています。

`login` 時に指定したサイトURLがデフォルトルートとして保存され、以降は `spo://相対パス` の形式で省略表記できます。

## セットアップ

### Entra IDアプリの設定

CLI用のアプリ登録を作成し、パブリッククライアントとして構成します。

1. アプリを登録（シングルテナントまたはマルチテナント）。
2. モバイルとデスクトップのプラットフォームを追加し、リダイレクトURIに `http://localhost` を設定。
3. パブリッククライアントフローを有効化。
4. SharePointの委任権限を追加:
	- `AllSites.Read`（`ls`のみの場合）
	- `AllSites.ReadWrite`（`cp`のアップロード/コピーに必要）
5. 必要に応じてテナントで管理者の同意を付与。

実行前に環境変数を設定します。

- `SPO_CLIENT_ID`: Entra IDアプリのクライアントID
- `SPO_TENANT_ID`: テナントID（省略可、既定: organizations）
- `SPO_SITE`: `login`の既定サイトURL（省略可）

## ビルド

```
dotnet build
```

## 実行

```
dotnet run -- login --site https://contoso.sharepoint.com/sites/demo/Shared%20Documents
```

```
dotnet run -- ls spo://
```

```
dotnet run -- ls spo://フォルダA
```

```
dotnet run -- cp spo://a.txt .\a.txt
```

```
dotnet run -- cp .\b.txt spo://フォルダB/b.txt
```

## 使い方の詳細

### ログインURLの正規化

`login` コマンドで `--site` に指定するURLは、フォルダやファイルの完全パスでも自動的にサイトルートに正規化されます。

以下のどの形式を指定しても、最終的には `https://contoso.sharepoint.com/sites/demo` がデフォルトルートとして保存されます：

```
# サイトルートで指定
dotnet run -- login --site https://contoso.sharepoint.com/sites/demo

# Shared Documents の階層で指定
dotnet run -- login --site https://contoso.sharepoint.com/sites/demo/Shared%20Documents/フォルダA

# ファイルのフルパスで指定
dotnet run -- login --site https://contoso.sharepoint.com/sites/demo/Shared%20Documents/フォルダA/ファイル.txt
```

### spo:// による短縮表記

`login` で保存されたデフォルトルートの **Shared Documents 配下** が、`spo://` のデフォルトになります。

```
# ログイン時のサイト: https://contoso.sharepoint.com/sites/demo

# 以下のコマンドは
dotnet run -- ls spo://フォルダA/フォルダB

# 実際には以下と同等です
dotnet run -- ls https://contoso.sharepoint.com/sites/demo/Shared%20Documents/フォルダA/フォルダB
```

### フォルダの一括ダウンロード

`cp` コマンドで転送先に末尾の `/` をつけると、フォルダ配下のファイルを一括ダウンロード（またはアップロード）できます：

```
# Shared Documents/フォルダA 配下のすべてをダウンロード
dotnet run -- cp spo://フォルダA/ .\ローカルフォルダ\

# ローカルフォルダ配下のすべてをアップロード
dotnet run -- cp .\ローカルフォルダ\ spo://フォルダA/
```

### 再帰ダウンロード（3階層まで）

`--recursive` フラグを指定すると、フォルダを最大3階層まで再帰的にダウンロードできます：

```
# Shared Documents/root 配下を最大3階層までダウンロード
# root/a.txt (1階層)
# root/dirB/b.txt (2階層)
# root/dirB/dirC/c.txt (3階層) ✅ ここまで
# root/dirB/dirC/dirD/d.txt (4階層) ❌ スキップ
dotnet run -- cp --recursive spo://root/ .\ローカルフォルダ\
```

フラグの位置は柔軟に対応しています：
- `cp --recursive spo://folder/ .\local\`
- `cp spo://folder/ .\local\ --recursive`

#### 注意事項

- **`--recursive` はダウンロード時のみ対応** - アップロード時に `--recursive` を指定しても無視されます
- **ダウンロード時にフォルダは自動作成** - ローカル側の対象ディレクトリが存在しない場合、自動的に作成されます
- **フォルダ指定時は末尾の `/` が必須** - ローカル側が入出力の向きを判定するため、必ず末尾に `/` をつけてください
