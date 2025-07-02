package main

import (
	"context"
	"fmt"
	"path"
	"strconv"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"
)

type GoogleSheets struct {
	cfg  *Config
	auth *GoogleAuth
}

func NewGoogleSheets(cfg *Config, auth *GoogleAuth) (*GoogleSheets, error) {
	return &GoogleSheets{
		cfg:  cfg,
		auth: auth,
	}, nil
}

type CopySheetRequest struct {
	SrcSpreadsheetName string `json:"src_spreadsheet"`
	SrcSheetName       string `json:"src_sheet"`
	DstSpreadsheetName string `json:"dst_spreadsheet"`
	DstSheetName       string `json:"dst_sheet"`
}

var CopySheetInputSchema = mcp.Input(
	mcp.Property("src_spreadsheet", mcp.Description("source spreadsheet name"), mcp.Required(true)),
	mcp.Property("src_sheet", mcp.Description("source sheet name"), mcp.Required(true)),
	mcp.Property("dst_spreadsheet", mcp.Description("destination spreadsheet name"), mcp.Required(true)),
	mcp.Property("dst_sheet", mcp.Description("destination sheet name"), mcp.Required(true)),
)

type RenameSheetRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	NewName         string `json:"new_name"`
}

var RenameSheetInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("new_name", mcp.Description("new sheet name"), mcp.Required(true)),
)

type ListSheetsRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
}

var ListSheetsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
)

type GetSheetDataRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	Range           string `json:"range"`
}

var GetSheetDataInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("range", mcp.Description("cell range (e.g. A1:C10, default: all data)")),
)

type AddRowsRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	Count           int64  `json:"count"`
	StartRow        int64  `json:"start_row"`
}

var AddRowsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("count", mcp.Description("number of rows to add"), mcp.Required(true)),
	mcp.Property("start_row", mcp.Description("row index to start adding (1-based)")),
)

type AddColumnsRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	Count           int64  `json:"count"`
	StartColumn     int64  `json:"start_column"`
}

var AddColumnsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("count", mcp.Description("number of columns to add"), mcp.Required(true)),
	mcp.Property("start_column", mcp.Description("column index to start adding (1-based)")),
)

// 行削除リクエスト
type DeleteRowsRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	Count           int64  `json:"count"`
	StartRow        int64  `json:"start_row"`
}

var DeleteRowsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("count", mcp.Description("number of rows to delete"), mcp.Required(true)),
	mcp.Property("start_row", mcp.Description("row index to start deleting (1-based)"), mcp.Required(true)),
)

// 列削除リクエスト
type DeleteColumnsRequest struct {
	SpreadsheetName string `json:"spreadsheet"`
	SheetName       string `json:"sheet"`
	Count           int64  `json:"count"`
	StartColumn     int64  `json:"start_column"`
}

var DeleteColumnsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("count", mcp.Description("number of columns to delete"), mcp.Required(true)),
	mcp.Property("start_column", mcp.Description("column index to start deleting (1-based)"), mcp.Required(true)),
)

// セル編集リクエスト
type UpdateCellsRequest struct {
	SpreadsheetName string          `json:"spreadsheet"`
	SheetName       string          `json:"sheet"`
	Range           string          `json:"range"`
	Data            [][]interface{} `json:"data"`
}

var UpdateCellsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("range", mcp.Description("cell range (e.g. A1:C10)"), mcp.Required(true)),
	mcp.Property("data", mcp.Description("2D array of cell values to update"), mcp.Required(true)),
)

// 複数範囲のセル編集リクエスト
type BatchUpdateCellsRequest struct {
	SpreadsheetName string                     `json:"spreadsheet"`
	SheetName       string                     `json:"sheet"`
	Ranges          map[string][][]interface{} `json:"ranges"`
}

var BatchUpdateCellsInputSchema = mcp.Input(
	mcp.Property("spreadsheet", mcp.Description("spreadsheet name"), mcp.Required(true)),
	mcp.Property("sheet", mcp.Description("sheet name"), mcp.Required(true)),
	mcp.Property("ranges", mcp.Description("map of range to 2D array of cell values (e.g. {'A1:B2': [[1, 2], [3, 4]], 'D5:E6': [[5, 6], [7, 8]]})"), mcp.Required(true)),
)

// スプレッドシート名からスプレッドシートIDを取得する
func (gs *GoogleSheets) getSpreadsheetId(spreadsheetName string) (string, error) {
	return gs.getSpreadsheetIdWithContext(context.Background(), spreadsheetName)
}

// コンテキスト付きでスプレッドシート名からスプレッドシートIDを取得する
func (gs *GoogleSheets) getSpreadsheetIdWithContext(ctx context.Context, spreadsheetName string) (string, error) {
	// パスの正規化と検証
	filePath := path.Clean(spreadsheetName)
	if strings.HasPrefix(filePath, "..") || strings.HasPrefix(filePath, "/") {
		return "", fmt.Errorf("invalid path: directory traversal is not allowed")
	}

	// ルートフォルダから開始
	parentID := gs.cfg.FolderID

	// パスが空の場合はエラー
	if filePath == "." || filePath == "" {
		return "", fmt.Errorf("spreadsheet name cannot be empty")
	}

	// パスを分割
	parts := strings.Split(filePath, "/")

	// 各パスの部分を順番に検索
	for i, part := range parts {
		isLast := i == len(parts)-1

		// 最後の部分（ファイル名）の場合はスプレッドシートタイプを指定
		var query string
		if isLast {
			query = fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false", parentID, part)
		} else {
			// フォルダの場合
			query = fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.folder' and trashed = false", parentID, part)
		}
		
		service, err := gs.auth.GetDriveService(ctx)
		if err != nil {
			return "", fmt.Errorf("failed to get drive service: %w", err)
		}

		fileList, err := service.Files.List().
			Q(query).
			SupportsAllDrives(true).
			IncludeItemsFromAllDrives(true).
			Fields("files(id, name, mimeType)").
			Do()
		if err != nil {
			return "", fmt.Errorf("failed to find file/folder: %w", err)
		}

		if len(fileList.Files) == 0 {
			if isLast {
				return "", fmt.Errorf("spreadsheet not found: %s", spreadsheetName)
			} else {
				return "", fmt.Errorf("folder not found: %s", strings.Join(parts[:i+1], "/"))
			}
		}

		// 次の親IDを設定（同名ファイルが複数ある場合は最初のものを使用）
		parentID = fileList.Files[0].Id
	}

	return parentID, nil
}

// シートIDを取得する
func (gs *GoogleSheets) getSheetId(spreadsheetId string, sheetName string) (int64, error) {
	return gs.getSheetIdWithContext(context.Background(), spreadsheetId, sheetName)
}

// コンテキスト付きでシートIDを取得する
func (gs *GoogleSheets) getSheetIdWithContext(ctx context.Context, spreadsheetId string, sheetName string) (int64, error) {
	// スプレッドシートの情報を取得
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return 0, fmt.Errorf("failed to get sheets service: %w", err)
	}

	spreadsheet, err := service.Spreadsheets.Get(spreadsheetId).Do()
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

func (gs *GoogleSheets) CopySheetHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[CopySheetRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// ソーススプレッドシートのIDを取得
	srcSpreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SrcSpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get source spreadsheet ID: %w", err)
	}

	// 宛先スプレッドシートのIDを取得
	dstSpreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.DstSpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get destination spreadsheet ID: %w", err)
	}

	srcSheetName := request.SrcSheetName
	dstSheetName := request.DstSheetName

	// ソースシートのIDを取得
	srcSheetId, err := gs.getSheetIdWithContext(ctx, srcSpreadsheetId, srcSheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get source sheet ID: %w", err)
	}

	// シートをコピー
	copySheetRequest := &sheets.CopySheetToAnotherSpreadsheetRequest{
		DestinationSpreadsheetId: dstSpreadsheetId,
	}

	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	copyResponse, err := service.Spreadsheets.Sheets.CopyTo(srcSpreadsheetId, srcSheetId, copySheetRequest).Do()
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
	service, err = gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(dstSpreadsheetId, updateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to rename copied sheet: %w", err)
	}

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully copied to sheet '%s' in spreadsheet '%s'",
					request.SrcSheetName, request.SrcSpreadsheetName,
					request.DstSheetName, request.DstSpreadsheetName),
			},
		},
	}, nil
}

func (gs *GoogleSheets) RenameSheetHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[RenameSheetRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 新しい名前が空でないことを確認
	if request.NewName == "" {
		return nil, fmt.Errorf("new sheet name cannot be empty")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
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
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetId, updateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to rename sheet: %w", err)
	}

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully renamed to '%s'",
					request.SheetName, request.SpreadsheetName, request.NewName),
			},
		},
	}, nil
}

func (gs *GoogleSheets) ListSheetsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[ListSheetsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシート名からスプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	// スプレッドシートの情報を取得
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	spreadsheet, err := service.Spreadsheets.Get(spreadsheetId).Do()
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
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: result.String(),
			},
		},
	}, nil
}

func (gs *GoogleSheets) AddRowsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[AddRowsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 追加する行数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	if request.StartRow <= 0 {
		return nil, fmt.Errorf("start_row must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// リクエストを作成
	var batchRequest *sheets.BatchUpdateSpreadsheetRequest

	batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				InsertDimension: &sheets.InsertDimensionRequest{
					Range: &sheets.DimensionRange{
						SheetId:    sheetId,
						Dimension:  "ROWS",
						StartIndex: request.StartRow - 1,
						EndIndex:   request.StartRow + request.Count - 1,
					},
					InheritFromBefore: false,
				},
			},
		},
	}

	// 行を追加
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to add rows: %w", err)
	}

	// 成功メッセージを作成
	var message string
	if request.StartRow > 0 {
		message = fmt.Sprintf("Successfully added %d rows at index %d in sheet '%s' of spreadsheet '%s'",
			request.Count, request.StartRow, request.SheetName, request.SpreadsheetName)
	} else {
		message = fmt.Sprintf("Successfully added %d rows at the end of sheet '%s' of spreadsheet '%s'",
			request.Count, request.SheetName, request.SpreadsheetName)
	}

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message,
			},
		},
	}, nil
}

func (gs *GoogleSheets) AddColumnsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[AddColumnsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 追加する列数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	if request.StartColumn <= 0 {
		return nil, fmt.Errorf("start_column must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// リクエストを作成
	var batchRequest *sheets.BatchUpdateSpreadsheetRequest

	batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				InsertDimension: &sheets.InsertDimensionRequest{
					Range: &sheets.DimensionRange{
						SheetId:    sheetId,
						Dimension:  "COLUMNS",
						StartIndex: request.StartColumn - 1,
						EndIndex:   request.StartColumn + request.Count - 1,
					},
					InheritFromBefore: false,
				},
			},
		},
	}

	// 列を追加
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to add columns: %w", err)
	}

	// 成功メッセージを作成
	var message string
	if request.StartColumn > 0 {
		message = fmt.Sprintf("Successfully added %d columns at index %d in sheet '%s' of spreadsheet '%s'",
			request.Count, request.StartColumn, request.SheetName, request.SpreadsheetName)
	} else {
		message = fmt.Sprintf("Successfully added %d columns at the end of sheet '%s' of spreadsheet '%s'",
			request.Count, request.SheetName, request.SpreadsheetName)
	}

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message,
			},
		},
	}, nil
}

// 2次元配列のデータを表形式の文字列に変換する関数
func formatTableData(startColumn, startRow int64, values [][]interface{}) string {

	maxWidth := 0
	for _, row := range values {
		if width := len(row); width > maxWidth {
			maxWidth = width
		}
	}
	var builder strings.Builder
	headers := make([]string, 0, maxWidth+1)
	borders := make([]string, 0, maxWidth+1)
	headers = append(headers, " ")
	borders = append(borders, "---")
	for i := 0; i < maxWidth; i++ {
		headers = append(headers, fmt.Sprintf("**%s**", columnIndexToLetter(startColumn+int64(i))))
		borders = append(borders, "---")
	}
	builder.WriteString("| " + strings.Join(headers, " | ") + "|\n")
	builder.WriteString("|" + strings.Join(borders, "|") + "|\n")

	for i, row := range values {
		rows := make([]string, maxWidth+1)
		rows[0] = fmt.Sprintf("**%d**", startRow+int64(i))
		for j, cell := range row {
			rows[j+1] = fmt.Sprintf("%v", cell)
		}
		builder.WriteString("| " + strings.Join(rows, " | ") + " |\n")
	}
	return builder.String()
}

// 列インデックス（0-based）をA1表記の列文字（A, B, C, ...）に変換する関数
func columnIndexToLetter(index int64) string {
	index = index - 1
	var result string
	for {
		remainder := index % 26
		result = string('A'+byte(remainder)) + result
		index = index/26 - 1
		if index < 0 {
			break
		}
	}
	return result
}

// 単一範囲のセル編集ハンドラー
func (gs *GoogleSheets) UpdateCellsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[UpdateCellsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 範囲が指定されていることを確認
	if request.Range == "" {
		return nil, fmt.Errorf("range must be specified")
	}

	// データが空でないことを確認
	if len(request.Data) == 0 {
		return nil, fmt.Errorf("data cannot be empty")
	}

	// 範囲を完全な形式に変換（シート名を含む）
	fullRange := fmt.Sprintf("%s!%s", sheetName, request.Range)

	// 変更前のデータを取得
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	prevData, err := service.Spreadsheets.Values.Get(spreadsheetId, fullRange).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get previous data: %w", err)
	}

	// 変更前のデータの行数と列数を計算
	prevRowCount := len(prevData.Values)
	prevColCount := 0
	for _, row := range prevData.Values {
		if colCount := len(row); colCount > prevColCount {
			prevColCount = colCount
		}
	}

	// 値を更新するリクエストを作成
	valueRange := &sheets.ValueRange{
		Range:  fullRange,
		Values: request.Data,
	}

	// 値を更新
	service, err = gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	updateResponse, err := service.Spreadsheets.Values.Update(
		spreadsheetId, fullRange, valueRange).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		return nil, fmt.Errorf("failed to update cells: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully updated %d cells in range '%s' of sheet '%s' in spreadsheet '%s'",
		updateResponse.UpdatedCells, request.Range, request.SheetName, request.SpreadsheetName)

	// 変更前のデータの情報をメッセージに含める
	message += fmt.Sprintf("\n\nPrevious data for range '%s' has been saved (%d rows x %d columns). To undo this change, you can use the previous data.",
		request.Range, prevRowCount, prevColCount)

	// 変更前のデータを表示用に整形
	col, row, err := startIndexFromRange(request.Range)
	if err != nil {
		return nil, fmt.Errorf("failed to parse range: %w", err)
	}
	prevDataStr := "\n\nPrevious data details:\n\n" + formatTableData(col, row, prevData.Values)

	// レスポンスを作成（変更前のデータを含める）
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message + prevDataStr,
			},
		},
	}, nil
}

func startIndexFromRange(rangeStr string) (int64, int64, error) {
	// 範囲が指定されていない場合はエラー
	if rangeStr == "" {
		return 0, 0, fmt.Errorf("range must be specified")
	}

	start := ""
	startEnd := strings.Split(rangeStr, ":")
	if len(startEnd) >= 1 {
		start = startEnd[0]
	} else {
		return 0, 0, fmt.Errorf("invalid range: %s", rangeStr)
	}

	// ABC123 のような形式を分割
	var colStr, rowStr string
	i := 0
	for _, c := range start {
		if c >= 'A' && c <= 'Z' {
			colStr += string(c)
			i++
		} else {
			break
		}
	}
	rowStr = start[i:]

	var col, row int64
	for _, c := range colStr {
		col = col*26 + int64(c-'A'+1)
	}
	row, err := strconv.ParseInt(rowStr, 10, 64)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid range: %s", rangeStr)
	}

	return col, row, nil
}

// 複数範囲のセル一括編集ハンドラー
func (gs *GoogleSheets) BatchUpdateCellsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[BatchUpdateCellsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 範囲が指定されていることを確認
	if len(request.Ranges) == 0 {
		return nil, fmt.Errorf("ranges cannot be empty")
	}

	// 変更前のデータを保存するマップ
	previousData := make(map[string][][]interface{})

	// バッチ更新用のデータを作成
	var data []*sheets.ValueRange
	for rangeStr, values := range request.Ranges {
		// データが空でないことを確認
		if len(values) == 0 {
			return nil, fmt.Errorf("data for range '%s' cannot be empty", rangeStr)
		}

		// 範囲を完全な形式に変換（シート名を含む）
		fullRange := fmt.Sprintf("%s!%s", sheetName, rangeStr)

		// 変更前のデータを取得
		service, err := gs.auth.GetSheetsService(ctx)
		if err != nil {
			return nil, fmt.Errorf("failed to get sheets service: %w", err)
		}

		prevData, err := service.Spreadsheets.Values.Get(spreadsheetId, fullRange).Do()
		if err != nil {
			return nil, fmt.Errorf("failed to get previous data for range '%s': %w", rangeStr, err)
		}
		previousData[rangeStr] = prevData.Values

		// ValueRangeを作成
		valueRange := &sheets.ValueRange{
			Range:  fullRange,
			Values: values,
		}

		data = append(data, valueRange)
	}

	// バッチ更新リクエストを作成
	batchUpdateRequest := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data:             data,
	}

	// バッチ更新を実行
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	batchUpdateResponse, err := service.Spreadsheets.Values.BatchUpdate(
		spreadsheetId, batchUpdateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to batch update cells: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully updated %d cells across %d ranges in sheet '%s' in spreadsheet '%s'",
		batchUpdateResponse.TotalUpdatedCells, batchUpdateResponse.TotalUpdatedSheets,
		request.SheetName, request.SpreadsheetName)

	// 変更前のデータの情報をメッセージに含める
	message += "\n\nPrevious data for the following ranges has been saved:"
	for rangeStr := range previousData {
		message += fmt.Sprintf("\n- %s", rangeStr)
	}
	message += "\n\nTo undo these changes, you can use the previous data."

	// 変更前のデータを表示用に整形
	var prevDataStr strings.Builder
	prevDataStr.WriteString("\n\nPrevious data details:\n\n")
	for rangeStr, values := range previousData {
		col, row, err := startIndexFromRange(rangeStr)
		if err != nil {
			return nil, fmt.Errorf("failed to parse range: %w", err)
		}
		prevDataStr.WriteString(formatTableData(col, row, values))
	}

	// レスポンスを作成（変更前のデータを含める）
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message + prevDataStr.String(),
			},
		},
	}, nil
}

func (gs *GoogleSheets) GetSheetDataHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[GetSheetDataRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 範囲が指定されていない場合はシート全体を取得
	range_ := sheetName
	if request.Range != "" {
		range_ = fmt.Sprintf("%s!%s", sheetName, request.Range)
	}

	// シートデータを取得
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	resp, err := service.Spreadsheets.Values.Get(spreadsheetId, range_).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet data: %w", err)
	}

	// 結果を整形
	var result strings.Builder
	result.WriteString(fmt.Sprintf("Data from sheet '%s' in spreadsheet '%s'",
		request.SheetName, request.SpreadsheetName))
	if request.Range != "" {
		result.WriteString(fmt.Sprintf(" (range: %s)", request.Range))
	}
	result.WriteString(":\n\n")

	// データがない場合
	if len(resp.Values) == 0 {
		result.WriteString("No data found.")
		return &mcp.CallToolResultFor[any]{
			Content: []mcp.Content{
				&mcp.TextContent{
					Text: result.String(),
				},
			},
		}, nil
	}

	// 各行のデータを表示
	var (
		startCol int64 = 1
		startRow int64 = 1
	)
	if request.Range != "" {
		startCol, startRow, err = startIndexFromRange(request.Range)
		if err != nil {
			return nil, fmt.Errorf("failed to parse range: %w", err)
		}
	}
	result.WriteString(formatTableData(startCol, startRow, resp.Values))

	// 行と列の数を表示
	rowCount := len(resp.Values)
	colCount := 0
	for _, row := range resp.Values {
		if count := len(row); count > colCount {
			colCount = count
		}
	}
	result.WriteString(fmt.Sprintf("\nTotal: %d rows x %d columns\n", rowCount, colCount))

	// 成功レスポンスを返す
	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: result.String(),
			},
		},
	}, nil
}

// 行削除ハンドラー
func (gs *GoogleSheets) DeleteRowsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[DeleteRowsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 削除する行数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	// 開始行が指定されていることを確認
	if request.StartRow <= 0 {
		return nil, fmt.Errorf("start row must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// 削除前のデータを取得（削除範囲のデータを保存）
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	// 削除する範囲を指定（A1表記に変換）
	startRowA1 := request.StartRow
	endRowA1 := startRowA1 + request.Count - 1
	rangeToDelete := fmt.Sprintf("%s!%d:%d", sheetName, startRowA1, endRowA1)

	// 削除前のデータを取得
	prevData, err := service.Spreadsheets.Values.Get(spreadsheetId, rangeToDelete).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get data before deletion: %w", err)
	}

	// 行を削除するリクエストを作成
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				DeleteDimension: &sheets.DeleteDimensionRequest{
					Range: &sheets.DimensionRange{
						SheetId:    sheetId,
						Dimension:  "ROWS",
						StartIndex: request.StartRow - 1,
						EndIndex:   request.StartRow + request.Count - 1,
					},
				},
			},
		},
	}

	// 行を削除
	service, err = gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to delete rows: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully deleted %d rows starting at index %d in sheet '%s' of spreadsheet '%s'",
		request.Count, request.StartRow, request.SheetName, request.SpreadsheetName)

	// 削除前のデータの情報をメッセージに含める
	prevRowCount := len(prevData.Values)
	prevColCount := 0
	if prevRowCount > 0 {
		prevColCount = len(prevData.Values[0])
	}

	message += fmt.Sprintf("\n\nDeleted data (%d rows x %d columns) has been saved. To undo this change, you can use the saved data.",
		prevRowCount, prevColCount)

	// 削除前のデータを表示用に整形

	prevDataStr := "\n\nDeleted data details:\n\n" + formatTableData(1, request.StartRow, prevData.Values)

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message + prevDataStr,
			},
		},
	}, nil
}

// 列削除ハンドラー
func (gs *GoogleSheets) DeleteColumnsHandler(ctx context.Context, cc *mcp.ServerSession, params *mcp.CallToolParamsFor[DeleteColumnsRequest]) (*mcp.CallToolResultFor[any], error) {
	request := params.Arguments
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetIdWithContext(ctx, request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 削除する列数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	// 開始列が指定されていることを確認
	if request.StartColumn <= 0 {
		return nil, fmt.Errorf("start column must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// 削除前のデータを取得（削除範囲のデータを保存）
	service, err := gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	// 削除する範囲を指定（A1表記に変換）
	startColA1 := columnIndexToLetter(request.StartColumn)
	endColA1 := columnIndexToLetter(request.StartColumn + request.Count - 1)
	rangeToDelete := fmt.Sprintf("%s!%s:%s", sheetName, startColA1, endColA1)

	// 削除前のデータを取得
	prevData, err := service.Spreadsheets.Values.Get(spreadsheetId, rangeToDelete).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get data before deletion: %w", err)
	}

	// 列を削除するリクエストを作成
	batchRequest := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			{
				DeleteDimension: &sheets.DeleteDimensionRequest{
					Range: &sheets.DimensionRange{
						SheetId:    sheetId,
						Dimension:  "COLUMNS",
						StartIndex: request.StartColumn - 1,
						EndIndex:   request.StartColumn + request.Count - 1,
					},
				},
			},
		},
	}

	// 列を削除
	service, err = gs.auth.GetSheetsService(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheets service: %w", err)
	}

	_, err = service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to delete columns: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully deleted %d columns starting at index %d in sheet '%s' of spreadsheet '%s'",
		request.Count, request.StartColumn, request.SheetName, request.SpreadsheetName)

	// 削除前のデータの情報をメッセージに含める
	prevRowCount := len(prevData.Values)
	prevColCount := 0
	for _, row := range prevData.Values {
		if width := len(row); width > prevColCount {
			prevColCount = width
		}
	}

	message += fmt.Sprintf("\n\nDeleted data (%d rows x %d columns) has been saved. To undo this change, you can use the saved data.",
		prevRowCount, prevColCount)

	// 削除前のデータを表示用に整形
	prevDataStr := "\n\nDeleted data details:\n\n" + formatTableData(request.StartColumn, 1, prevData.Values)

	return &mcp.CallToolResultFor[any]{
		Content: []mcp.Content{
			&mcp.TextContent{
				Text: message + prevDataStr,
			},
		},
	}, nil
}
