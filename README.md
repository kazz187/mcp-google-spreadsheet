# MCP Google Spreadsheet

MCP (Model Context Protocol) サーバーとして実装された Google Spreadsheet および Google Drive 操作ツールです。このツールを使用することで、AI アシスタントが Google Spreadsheet や Google Drive のファイルを操作できるようになります。

## 機能

### Google Drive 操作

- **list_files**: Google Drive のファイル一覧を取得
- **copy_file**: Google Drive のファイルをコピー
- **rename_file**: Google Drive のファイル名を変更

### Google Spreadsheet 操作

- **list_sheets**: スプレッドシート内のシート一覧を取得
- **copy_sheet**: スプレッドシート内のシートをコピー
- **rename_sheet**: スプレッドシート内のシート名を変更
- **get_sheet_data**: シートのデータを取得
- **add_rows**: シートに行を追加
- **add_columns**: シートに列を追加
- **update_cells**: 単一範囲のセルを更新
- **batch_update_cells**: 複数範囲のセルを一括更新

## 前提条件

- Go 1.24 以上
- Google Cloud Platform のプロジェクトと API 有効化
  - Google Drive API
  - Google Sheets API

## インストール

```bash
go install github.com/kazz187/mcp-google-spreadsheet@latest
```

これにより、`$GOPATH/bin` ディレクトリに `mcp-google-spreadsheet` バイナリがインストールされます。

## 設定

以下の環境変数を設定する必要があります：

- `GSUITE_CLIENT_SECRET_PATH`: Google API のクライアントシークレットファイルのパス
- `GSUITE_TOKEN_PATH`: Google API のトークンファイルのパス（存在しない場合は自動的に作成されます）
- `GSUITE_FOLDER_ID`: 操作対象とする Google Drive のフォルダ ID

### Google API の設定手順

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを作成
3. Google Drive API と Google Sheets API を有効化
4. 認証情報を作成（OAuth クライアント ID）
5. クライアントシークレットをダウンロード

## 使用方法

### 起動

```bash
export GSUITE_CLIENT_SECRET_PATH=/path/to/client_secret.json
export GSUITE_TOKEN_PATH=/path/to/token.json
export GSUITE_FOLDER_ID=your_folder_id
mcp-google-spreadsheet
```

`go install` でインストールした場合は、`$GOPATH/bin` が PATH に含まれていることを確認してください。

初回起動時は認証が必要です。表示される URL をブラウザで開き、アクセスを許可した後、認証コードを入力してください。

### MCP 設定

Claude や ChatGPT などの AI アシスタントで使用するには、MCP の設定ファイルに以下のように追加します：

```json
{
  "mcpServers": {
    "mcp_google_spreadsheet": {
      "command": "mcp-google-spreadsheet",
      "args": [],
      "env": {
        "GSUITE_CLIENT_SECRET_PATH": "/path/to/client_secret.json",
        "GSUITE_TOKEN_PATH": "/path/to/token.json",
        "GSUITE_FOLDER_ID": "your_folder_id"
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
