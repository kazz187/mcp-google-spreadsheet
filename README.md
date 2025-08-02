# MCP Google Spreadsheet

MCP (Model Context Protocol) サーバーとして実装された Google Spreadsheet および Google Drive 操作ツールです。このツールを使用することで、AI アシスタント（Claude Desktop など）が Google Spreadsheet や Google Drive のファイルを操作できるようになります。

## 機能

### Google Drive 操作

- **google_drive_list_files**: Google Drive 内のファイルとフォルダを一覧表示
- **google_drive_copy_file**: ファイルまたはフォルダを別の場所にコピー
- **google_drive_rename_file**: ファイルまたはフォルダの名前を変更

### Google Spreadsheet 操作

- **google_sheets_list_sheets**: スプレッドシート内のシート（タブ）一覧を取得
- **google_sheets_copy_sheet**: シートを別のスプレッドシートにコピー
- **google_sheets_rename_sheet**: シートの名前を変更
- **google_sheets_read_data**: シートのデータを読み取り（スプレッドシートを「開く」操作）
- **google_sheets_add_rows**: シートに空の行を挿入
- **google_sheets_add_columns**: シートに空の列を挿入
- **google_sheets_update_cells**: 指定範囲のセルの値を更新
- **google_sheets_batch_update_cells**: 複数範囲のセルを一括更新
- **google_sheets_delete_rows**: シートから行を削除
- **google_sheets_delete_columns**: シートから列を削除

## 使用ワークフロー

1. `google_drive_list_files` で Google Drive 内のファイルを探索し、スプレッドシートを見つける
2. `google_sheets_list_sheets` で特定のスプレッドシート内のシート一覧を確認
3. `google_sheets_read_data` でシートの内容を表示
4. 必要に応じて他の `google_sheets_*` ツールでデータを編集

## 前提条件

- Go 1.24 以上
- Google Cloud Platform のプロジェクトと API 有効化
  - Google Drive API
  - Google Sheets API
- Google OAuth 2.0 認証の設定

## インストール

```bash
go install github.com/kazz187/mcp-google-spreadsheet@latest
```

これにより、`$GOPATH/bin` ディレクトリに `mcp-google-spreadsheet` バイナリがインストールされます。

## 設定

以下の環境変数を設定する必要があります：

- `MCPGS_CLIENT_SECRET_PATH`: Google API のクライアントシークレットファイルのパス (https://developers.google.com/identity/protocols/oauth2/native-app?hl=ja)
- `MCPGS_TOKEN_PATH`: Google API のトークンファイルのパス（存在しない場合は自動的に作成されます）
- `MCPGS_FOLDER_ID`: 操作対象とする Google Drive のフォルダ ID（フォルダを右クリック → リンクを取得 → URLの最後の部分）

### Google API の設定手順

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを作成
3. Google Drive API と Google Sheets API を有効化
4. 認証情報を作成（OAuth クライアント ID）
   - アプリケーションの種類：「デスクトップアプリケーション」を選択
5. クライアントシークレット JSON をダウンロード

## 使用方法

### 起動

```bash
export MCPGS_CLIENT_SECRET_PATH=/path/to/client_secret.json
export MCPGS_TOKEN_PATH=/path/to/token.json
export MCPGS_FOLDER_ID=your_folder_id
mcp-google-spreadsheet
```

`go install` でインストールした場合は、`$GOPATH/bin` が PATH に含まれていることを確認してください。

初回起動時は認証が必要です。ブラウザが自動的に開き、Google アカウントでの認証画面が表示されます。認証が完了すると自動的にアプリケーションに戻ります。ブラウザが自動的に開かない場合は、コンソールに表示される URL をブラウザで開いてください。

### Claude Desktop での設定

Claude Desktop で使用する場合は、設定ファイル（macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`）に以下のように追加します：

```json
{
  "mcpServers": {
    "mcp_google_spreadsheet": {
      "command": "mcp-google-spreadsheet",
      "args": [],
      "env": {
        "MCPGS_CLIENT_SECRET_PATH": "/path/to/client_secret.json",
        "MCPGS_TOKEN_PATH": "/path/to/token.json",
        "MCPGS_FOLDER_ID": "your_folder_id"
      }
    }
  }
}
```

`go install` でインストールした場合は、`command` に絶対パスを指定する代わりに、上記のように実行ファイル名のみを指定することもできます。その場合は、MCP サーバーを実行するユーザーの PATH に `$GOPATH/bin` が含まれていることを確認してください。

## セキュリティ

- 指定されたフォルダ ID 内のファイルのみにアクセスが制限されます
- ディレクトリトラバーサル攻撃（`../` などを使用したパス指定）は防止されます
- ユーザーから指定されたファイルが指定フォルダ内に存在するかが検証されます

