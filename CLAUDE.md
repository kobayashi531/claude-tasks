# Claude Tasks

## freeeファイルボックスアップロード

freee MCPはファイルアップロード非対応のため、PowerShellスクリプトを使用する。

### 実行方法

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\kobayashi\Desktop\Claude\freee_upload.ps1"
```

### 手順

1. スクリプト実行 → ブラウザでfreee認証画面が開く
2. freeeにログインして「許可する」をクリック
3. ブラウザに「Auth OK! You can close this tab.」と表示されれば認証成功
4. 自動でファイルボックスにアップロードされる

### 設定値

- スクリプト: `C:\Users\kobayashi\Desktop\Claude\freee_upload.ps1`
- Company ID: 787791（株式会社シー・コネクト）
- Redirect URI: `http://localhost:8080/callback`（freee開発者コンソールに登録済み）
- Client ID: 709292942535479

### 別ファイルをアップロードする場合

スクリプト内の `$FilePath` と `Content-Type` を変更して実行する。
