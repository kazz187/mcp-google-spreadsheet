package main

import (
	"context"
	"fmt"
	"path"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/drive/v3"
)

type GoogleDrive struct {
	cfg  *Config
	auth *GoogleAuth
}

func NewGoogleDrive(cfg *Config, auth *GoogleAuth) (*GoogleDrive, error) {
	return &GoogleDrive{
		cfg:  cfg,
		auth: auth,
	}, nil
}

type ListFilesRequest struct {
	Path string `json:"path"`
}

var ListFilesInputSchema = mcp.Input(
	mcp.Property("path", mcp.Description("directory path (default: root directory)")),
)

type CopyFileRequest struct {
	SrcPath string `json:"src_path"`
	DstPath string `json:"dst_path"`
}

var CopyFileInputSchema = mcp.Input(
	mcp.Property("src_path", mcp.Description("source path"), mcp.Required(true)),
	mcp.Property("dst_path", mcp.Description("destination path"), mcp.Required(true)),
)

type RenameFileRequest struct {
	Path    string `json:"path"`
	NewName string `json:"new_name"`
}

var RenameFileInputSchema = mcp.Input(
	mcp.Property("path", mcp.Description("file path"), mcp.Required(true)),
	mcp.Property("new_name", mcp.Description("new file name"), mcp.Required(true)),
)

// パスからファイルIDを取得する
func (gd *GoogleDrive) getFileIDByPath(ctx context.Context, filePath string) (string, error) {
	return gd.getFileIDByPathWithContext(ctx, filePath)
}

// コンテキスト付きでパスからファイルIDを取得する
func (gd *GoogleDrive) getFileIDByPathWithContext(ctx context.Context, filePath string) (string, error) {
	// パスの正規化と検証
	filePath = path.Clean(filePath)
	if strings.HasPrefix(filePath, "..") || strings.HasPrefix(filePath, "/") {
		return "", fmt.Errorf("invalid path: directory traversal is not allowed")
	}

	// ルートフォルダから開始
	parentID := gd.cfg.FolderID

	// パスが空の場合はルートフォルダを返す
	if filePath == "." || filePath == "" {
		return parentID, nil
	}

	// パスを分割
	parts := strings.Split(filePath, "/")

	// 各パスの部分を順番に検索
	for i, part := range parts {
		isLast := i == len(parts)-1

		// 現在のフォルダ内のファイル/フォルダを検索
		query := fmt.Sprintf("'%s' in parents and name = '%s' and trashed = false", parentID, part)
		service, err := gd.auth.GetDriveService(ctx)
		if err != nil {
			return "", fmt.Errorf("failed to get drive service: %w", err)
		}

		fileList, err := service.Files.List().
			Q(query).
			SupportsAllDrives(true).         // 共有ドライブ対応
			IncludeItemsFromAllDrives(true). // 共有ドライブ対応
			Fields("files(id, mimeType)").
			Do()
		if err != nil {
			return "", fmt.Errorf("failed to list files: %w", err)
		}

		if len(fileList.Files) == 0 {
			// 最後のパス部分で、ファイルが存在しない場合はエラー
			if isLast {
				return "", fmt.Errorf("file not found: %s", filePath)
			}
			return "", fmt.Errorf("path not found: %s", strings.Join(parts[:i+1], "/"))
		}

		// 次の親IDを設定（同名ファイルが複数ある場合は最初のものを使用）
		parentID = fileList.Files[0].Id
	}

	return parentID, nil
}

// パスからファイルの親フォルダIDとファイル名を取得する
func (gd *GoogleDrive) getParentIDAndFileName(ctx context.Context, filePath string) (string, string, error) {
	// パスの正規化と検証
	filePath = path.Clean(filePath)
	if strings.HasPrefix(filePath, "..") || strings.HasPrefix(filePath, "/") {
		return "", "", fmt.Errorf("invalid path: directory traversal is not allowed")
	}

	// パスが空の場合はエラー
	if filePath == "." || filePath == "" {
		return "", "", fmt.Errorf("invalid path: path is empty")
	}

	// パスを親ディレクトリとファイル名に分割
	dir, fileName := path.Split(filePath)

	// 親ディレクトリのIDを取得
	parentID := gd.cfg.FolderID
	if dir != "" {
		// 末尾のスラッシュを削除
		dir = strings.TrimSuffix(dir, "/")
		var err error
		parentID, err = gd.getFileIDByPath(ctx, dir)
		if err != nil {
			return "", "", err
		}
	}

	return parentID, fileName, nil
}

func (gd *GoogleDrive) ListFilesHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[ListFilesRequest]) (*mcp.CallToolResultFor[any], error) {
	// パスが指定されていない場合はルートディレクトリを使用
	dirPath := params.Arguments.Path
	if dirPath == "" {
		dirPath = "."
	}

	// 指定されたパスのフォルダIDを取得
	folderID, err := gd.getFileIDByPathWithContext(ctx, dirPath)
	if err != nil {
		return nil, fmt.Errorf("failed to get folder ID: %w", err)
	}

	// フォルダ内のファイルとフォルダを取得
	query := fmt.Sprintf("'%s' in parents and trashed = false", folderID)
	service, err := gd.auth.GetDriveService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get drive service: %w", err)
	}

	fileList, err := service.Files.List().
		Q(query).
		SupportsAllDrives(true).         // 共有ドライブ対応
		IncludeItemsFromAllDrives(true). // 共有ドライブ対応
		Fields("files(id, name, mimeType, createdTime, modifiedTime, size)").
		OrderBy("name").
		Do()
	if err != nil {
		return nil, fmt.Errorf("failed to list files: %w", err)
	}

	// 結果を整形
	var result strings.Builder
	result.WriteString(fmt.Sprintf("Files in directory '%s':\n\n", dirPath))

	// フォルダを先に表示
	result.WriteString("Folders:\n")
	folderCount := 0
	for _, file := range fileList.Files {
		if file.MimeType == "application/vnd.google-apps.folder" {
			result.WriteString(fmt.Sprintf("- %s (Folder)\n", file.Name))
			folderCount++
		}
	}
	if folderCount == 0 {
		result.WriteString("  No folders found\n")
	}

	// ファイルを表示
	result.WriteString("\nFiles:\n")
	fileCount := 0
	for _, file := range fileList.Files {
		if file.MimeType != "application/vnd.google-apps.folder" {
			// Google Docsなどの特殊なファイルタイプを判別
			fileType := "File"
			switch file.MimeType {
			case "application/vnd.google-apps.document":
				fileType = "Google Doc"
			case "application/vnd.google-apps.spreadsheet":
				fileType = "Google Spreadsheet"
			case "application/vnd.google-apps.presentation":
				fileType = "Google Presentation"
			case "application/vnd.google-apps.form":
				fileType = "Google Form"
			}

			// ファイルサイズ（Google Docsなどは表示されない）
			sizeInfo := ""
			if file.Size > 0 {
				sizeInfo = fmt.Sprintf(", Size: %d bytes", file.Size)
			}

			result.WriteString(fmt.Sprintf("- %s (%s%s)\n", file.Name, fileType, sizeInfo))
			fileCount++
		}
	}
	if fileCount == 0 {
		result.WriteString("  No files found\n")
	}

	// 合計数
	result.WriteString(fmt.Sprintf("\nTotal: %d folders, %d files\n", folderCount, fileCount))

	// 成功レスポンスを返す
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{&mcp.TextContent{Text: result.String()}},
	}, nil
}

func (gd *GoogleDrive) CopyFileHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[CopyFileRequest]) (*mcp.CallToolResultFor[any], error) {
	// ソースファイルのIDを取得
	srcFileID, err := gd.getFileIDByPathWithContext(ctx, params.Arguments.SrcPath)
	if err != nil {
		return nil, fmt.Errorf("failed to get source file ID: %w", err)
	}

	// ソースファイルの情報を取得（共有ドライブ対応のため supportsAllDrives を追加）
	service, err := gd.auth.GetDriveService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get drive service: %w", err)
	}

	srcFile, err := service.Files.Get(srcFileID).
		SupportsAllDrives(true).
		Fields("name", "mimeType").
		Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get source file: %w", err)
	}

	// 宛先の親フォルダIDとファイル名を取得
	dstParentID, dstFileName, err := gd.getParentIDAndFileName(ctx, params.Arguments.DstPath)
	if err != nil {
		return nil, fmt.Errorf("failed to get destination parent ID and file name: %w", err)
	}

	// ファイル名が指定されていない場合は、ソースファイルと同じ名前を使用
	if dstFileName == "" {
		dstFileName = srcFile.Name
	}

	// ファイルをコピー
	copiedFile := &drive.File{
		Name:    dstFileName,
		Parents: []string{dstParentID},
	}

	// サービスを再取得（前のサービスが期限切れの可能性があるため）
	service, err = gd.auth.GetDriveService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get drive service: %w", err)
	}

	result, err := service.Files.Copy(srcFileID, copiedFile).
		SupportsAllDrives(true).
		Do()
	if err != nil {
		return nil, fmt.Errorf("failed to copy file: %w", err)
	}

	// 成功レスポンスを返す
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{&mcp.TextContent{Text: fmt.Sprintf("File copied successfully. New file ID: %s", result.Id)}},
	}, nil
}

func (gd *GoogleDrive) RenameFileHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[RenameFileRequest]) (*mcp.CallToolResultFor[any], error) {
	// ファイルのIDを取得
	fileID, err := gd.getFileIDByPathWithContext(ctx, params.Arguments.Path)
	if err != nil {
		return nil, fmt.Errorf("failed to get file ID: %w", err)
	}

	// ファイルの情報を取得して存在確認（共有ドライブ対応）
	service, err := gd.auth.GetDriveService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get drive service: %w", err)
	}

	_, err = service.Files.Get(fileID).
		SupportsAllDrives(true).
		Fields("name").
		Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get file: %w", err)
	}

	// 新しい名前が空でないことを確認
	if params.Arguments.NewName == "" {
		return nil, fmt.Errorf("new file name cannot be empty")
	}

	// ファイル名を更新
	updateFile := &drive.File{
		Name: params.Arguments.NewName,
	}

	// ファイル名を変更（共有ドライブ対応）
	// サービスを再取得（前のサービスが期限切れの可能性があるため）
	service, err = gd.auth.GetDriveService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get drive service: %w", err)
	}

	result, err := service.Files.Update(fileID, updateFile).
		SupportsAllDrives(true).
		Do()
	if err != nil {
		return nil, fmt.Errorf("failed to rename file: %w", err)
	}

	// 成功レスポンスを返す
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{&mcp.TextContent{Text: fmt.Sprintf("File renamed successfully to '%s'", result.Name)}},
	}, nil
}
