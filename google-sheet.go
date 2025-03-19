package main

import (
	"context"
	"fmt"
	"net/http"
	"strings"

	mcp "github.com/metoro-io/mcp-golang"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type GoogleSheets struct {
	cfg     *Config
	service *sheets.Service
	drive   *drive.Service
}

func NewGoogleSheets(ctx context.Context, cfg *Config, cli *http.Client) (*GoogleSheets, error) {
	service, err := sheets.NewService(ctx, option.WithHTTPClient(cli))
	if err != nil {
		return nil, fmt.Errorf("failed to create sheets service: %w", err)
	}

	driveService, err := drive.NewService(ctx, option.WithHTTPClient(cli))
	if err != nil {
		return nil, fmt.Errorf("failed to create drive service for sheets: %w", err)
	}

	return &GoogleSheets{
		cfg:     cfg,
		service: service,
		drive:   driveService,
	}, nil
}

type CopySheetRequest struct {
	SrcName string `json:"src_path" jsonschema:"required,description=source sheet name"`
	DstName string `json:"dst_path" jsonschema:"required,description=destination sheet name"`
}

type RenameSheetRequest struct {
	Path    string `json:"path" jsonschema:"required,description=sheet path"`
	NewName string `json:"new_name" jsonschema:"required,description=new sheet name"`
}

type ListSheetsRequest struct {
	SpreadsheetName string `json:"path" jsonschema:"required,description=spreadsheet name"`
}

type GetSheetDataRequest struct {
	Path  string `json:"path" jsonschema:"required,description=sheet path (format: SpreadsheetName/SheetName)"`
	Range string `json:"range" jsonschema:"description=cell range (e.g. A1:C10, default: all data)"`
}

// パスからスプレッドシートIDとシート名を抽出する
// 例: "MySpreadsheet/Sheet1" -> "spreadsheetId", "Sheet1"
func (gs *GoogleSheets) parseSheetPath(sheetPath string) (string, string, error) {
	parts := strings.Split(sheetPath, "/")
	if len(parts) != 2 {
		return "", "", fmt.Errorf("invalid sheet path format: %s (expected format: 'SpreadsheetName/SheetName')", sheetPath)
	}

	spreadsheetName := parts[0]
	sheetName := parts[1]

	// スプレッドシート名からIDを検索
	query := fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false", gs.cfg.FolderID, spreadsheetName)
	fileList, err := gs.drive.Files.List().Q(query).Fields("files(id, name)").Do()
	if err != nil {
		return "", "", fmt.Errorf("failed to find spreadsheet: %w", err)
	}

	if len(fileList.Files) == 0 {
		return "", "", fmt.Errorf("spreadsheet not found: %s", spreadsheetName)
	}

	spreadsheetId := fileList.Files[0].Id
	return spreadsheetId, sheetName, nil
}

// スプレッドシート名からスプレッドシートIDを取得する
func (gs *GoogleSheets) getSpreadsheetId(spreadsheetName string) (string, error) {
	// スプレッドシート名からIDを検索
	query := fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false", gs.cfg.FolderID, spreadsheetName)
	fileList, err := gs.drive.Files.List().Q(query).Fields("files(id, name)").Do()
	if err != nil {
		return "", fmt.Errorf("failed to find spreadsheet: %w", err)
	}

	if len(fileList.Files) == 0 {
		return "", fmt.Errorf("spreadsheet not found: %s", spreadsheetName)
	}

	return fileList.Files[0].Id, nil
}

// シートIDを取得する
func (gs *GoogleSheets) getSheetId(spreadsheetId string, sheetName string) (int64, error) {
	// スプレッドシートの情報を取得
	spreadsheet, err := gs.service.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		return 0, fmt.Errorf("failed to get spreadsheet: %w", err)
	}

	// シート名からシートIDを検索
	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}

	return 0, fmt.Errorf("sheet not found: %s", sheetName)
}

func (gs *GoogleSheets) CopySheetHandler(request CopySheetRequest) (*mcp.ToolResponse, error) {
	// ソースシートのパスを解析
	srcSpreadsheetId, srcSheetName, err := gs.parseSheetPath(request.SrcName)
	if err != nil {
		return nil, fmt.Errorf("failed to parse source sheet path: %w", err)
	}

	// 宛先シートのパスを解析
	dstSpreadsheetId, dstSheetName, err := gs.parseSheetPath(request.DstName)
	if err != nil {
		return nil, fmt.Errorf("failed to parse destination sheet path: %w", err)
	}

	// ソースシートのIDを取得
	srcSheetId, err := gs.getSheetId(srcSpreadsheetId, srcSheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get source sheet ID: %w", err)
	}

	// シートをコピー
	copySheetRequest := &sheets.CopySheetToAnotherSpreadsheetRequest{
		DestinationSpreadsheetId: dstSpreadsheetId,
	}

	copyResponse, err := gs.service.Spreadsheets.Sheets.CopyTo(srcSpreadsheetId, srcSheetId, copySheetRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to copy sheet: %w", err)
	}

	// コピーしたシートの名前を変更（指定された名前に）
	// 新しいシートのIDを取得
	newSheetId := copyResponse.SheetId

	// シート名を更新するリクエストを作成
	updateRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
					Properties: &sheets.SheetProperties{
						SheetId: newSheetId,
						Title:   dstSheetName,
					},
					Fields: "title",
				},
			},
		},
	}

	// シート名を更新
	_, err = gs.service.Spreadsheets.BatchUpdate(dstSpreadsheetId, updateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to rename copied sheet: %w", err)
	}

	return mcp.NewToolResponse(
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' successfully copied to '%s'", request.SrcName, request.DstName)),
	), nil
}

func (gs *GoogleSheets) RenameSheetHandler(request RenameSheetRequest) (*mcp.ToolResponse, error) {
	// シートのパスを解析
	spreadsheetId, sheetName, err := gs.parseSheetPath(request.Path)
	if err != nil {
		return nil, fmt.Errorf("failed to parse sheet path: %w", err)
	}

	// 新しい名前が空でないことを確認
	if request.NewName == "" {
		return nil, fmt.Errorf("new sheet name cannot be empty")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetId(spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// シート名を更新するリクエストを作成
	updateRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
					Properties: &sheets.SheetProperties{
						SheetId: sheetId,
						Title:   request.NewName,
					},
					Fields: "title",
				},
			},
		},
	}

	// シート名を更新
	_, err = gs.service.Spreadsheets.BatchUpdate(spreadsheetId, updateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to rename sheet: %w", err)
	}

	return mcp.NewToolResponse(
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' successfully renamed to '%s'", request.Path, request.NewName)),
	), nil
}

func (gs *GoogleSheets) ListSheetsHandler(request ListSheetsRequest) (*mcp.ToolResponse, error) {
	// スプレッドシート名からスプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	// スプレッドシートの情報を取得
	spreadsheet, err := gs.service.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet: %w", err)
	}

	// 結果を整形
	var result strings.Builder
	result.WriteString(fmt.Sprintf("Sheets in spreadsheet '%s':\n\n", request.SpreadsheetName))

	// シート情報を表示
	for i, sheet := range spreadsheet.Sheets {
		result.WriteString(fmt.Sprintf("%d. %s (ID: %d)\n", i+1, sheet.Properties.Title, sheet.Properties.SheetId))
	}

	// 合計数
	result.WriteString(fmt.Sprintf("\nTotal: %d sheets\n", len(spreadsheet.Sheets)))

	// 成功レスポンスを返す
	return mcp.NewToolResponse(
		mcp.NewTextContent(result.String()),
	), nil
}

func (gs *GoogleSheets) GetSheetDataHandler(request GetSheetDataRequest) (*mcp.ToolResponse, error) {
	// シートのパスを解析
	spreadsheetId, sheetName, err := gs.parseSheetPath(request.Path)
	if err != nil {
		return nil, fmt.Errorf("failed to parse sheet path: %w", err)
	}

	// 範囲が指定されていない場合はシート全体を取得
	range_ := sheetName
	if request.Range != "" {
		range_ = fmt.Sprintf("%s!%s", sheetName, request.Range)
	}

	// シートデータを取得
	resp, err := gs.service.Spreadsheets.Values.Get(spreadsheetId, range_).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet data: %w", err)
	}

	// 結果を整形
	var result strings.Builder
	result.WriteString(fmt.Sprintf("Data from sheet '%s'", request.Path))
	if request.Range != "" {
		result.WriteString(fmt.Sprintf(" (range: %s)", request.Range))
	}
	result.WriteString(":\n\n")

	// データがない場合
	if len(resp.Values) == 0 {
		result.WriteString("No data found.")
		return mcp.NewToolResponse(
			mcp.NewTextContent(result.String()),
		), nil
	}

	// 各行のデータを表示
	for i, row := range resp.Values {
		result.WriteString(fmt.Sprintf("Row %d: ", i+1))
		for j, cell := range row {
			if j > 0 {
				result.WriteString(" | ")
			}
			result.WriteString(fmt.Sprintf("%v", cell))
		}
		result.WriteString("\n")
	}

	// 行と列の数を表示
	rowCount := len(resp.Values)
	colCount := 0
	if rowCount > 0 {
		colCount = len(resp.Values[0])
	}
	result.WriteString(fmt.Sprintf("\nTotal: %d rows x %d columns\n", rowCount, colCount))

	// 成功レスポンスを返す
	return mcp.NewToolResponse(
		mcp.NewTextContent(result.String()),
	), nil
}
