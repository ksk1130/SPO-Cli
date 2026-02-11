# spo-cli

SharePoint Online のファイル操作をCLIで(aws s3コマンド風に)実行できるツールです。

## コマンド

- `login [--mfa]`
- `ls <site-or-folder-url>`
- `cp <from> <to>`

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
