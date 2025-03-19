package main

import (
	"context"
	"fmt"
	"net/http"
	"path"
	"strings"

	mcp "github.com/metoro-io/mcp-golang"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

type GoogleDrive struct {
	cfg     *Config
	service *drive.Service
}

func NewGoogleDrive(ctx context.Context, cfg *Config, client *http.Client) (*GoogleDrive, error) {
	service, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("failed to create drive service: %w", err)
	}
	return &GoogleDrive{
		cfg:     cfg,
		service: service,
	}, nil
}

type CopyFileRequest struct {
	SrcPath string `json:"src_path" jsonschema:"required,description=source path"`
	DstPath string `json:"dst_path" jsonschema:"required,description=destination path"`
}

// パスからファイルIDを取得する
func (gd *GoogleDrive) getFileIDByPath(filePath string) (string, error) {
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
		fileList, err := gd.service.Files.List().Q(query).Fields("files(id, mimeType)").Do()
		if err != nil {
			return "", fmt.Errorf("failed to list files: %w", err)
		}

		if len(fileList.Files) == 0 {
			// 最後のパス部分で、ファイルが存在しない場合は新規作成の可能性があるため、親フォルダIDを返す
			if isLast {
				return "", fmt.Errorf("file not found: %s", filePath)
			}
			return "", fmt.Errorf("path not found: %s", strings.Join(parts[:i+1], "/"))
		}

		// 次の親IDを設定
		parentID = fileList.Files[0].Id
	}

	return parentID, nil
}

// パスからファイルの親フォルダIDとファイル名を取得する
func (gd *GoogleDrive) getParentIDAndFileName(filePath string) (string, string, error) {
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
		parentID, err = gd.getFileIDByPath(dir)
		if err != nil {
			return "", "", err
		}
	}

	return parentID, fileName, nil
}

func (gd *GoogleDrive) CopyFileHandler(request CopyFileRequest) (*mcp.ToolResponse, error) {
	// ソースファイルのIDを取得
	srcFileID, err := gd.getFileIDByPath(request.SrcPath)
	if err != nil {
		return nil, fmt.Errorf("failed to get source file ID: %w", err)
	}

	// ソースファイルの情報を取得
	srcFile, err := gd.service.Files.Get(srcFileID).Fields("name", "mimeType").Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get source file: %w", err)
	}

	// 宛先の親フォルダIDとファイル名を取得
	dstParentID, dstFileName, err := gd.getParentIDAndFileName(request.DstPath)
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

	result, err := gd.service.Files.Copy(srcFileID, copiedFile).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to copy file: %w", err)
	}

	// 成功レスポンスを返す
	return mcp.NewToolResponse(
		mcp.NewTextContent(fmt.Sprintf("File copied successfully. New file ID: %s", result.Id)),
	), nil
}
