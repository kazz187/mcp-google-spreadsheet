package main

import (
	"context"
	"fmt"
	"strings"

	mcp "github.com/metoro-io/mcp-golang"
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
	SrcSpreadsheetName string `json:"src_spreadsheet" jsonschema:"required,description=source spreadsheet name"`
	SrcSheetName       string `json:"src_sheet" jsonschema:"required,description=source sheet name"`
	DstSpreadsheetName string `json:"dst_spreadsheet" jsonschema:"required,description=destination spreadsheet name"`
	DstSheetName       string `json:"dst_sheet" jsonschema:"required,description=destination sheet name"`
}

type RenameSheetRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	NewName         string `json:"new_name" jsonschema:"required,description=new sheet name"`
}

type ListSheetsRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
}

type GetSheetDataRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	Range           string `json:"range" jsonschema:"description=cell range (e.g. A1:C10, default: all data)"`
}

type AddRowsRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	Count           int64  `json:"count" jsonschema:"required,description=number of rows to add"`
	StartRow        int64  `json:"start_row" jsonschema:"description=row index to start adding (0-based, default: append to end)"`
}

type AddColumnsRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	Count           int64  `json:"count" jsonschema:"required,description=number of columns to add"`
	StartColumn     int64  `json:"start_column" jsonschema:"description=column index to start adding (0-based, default: append to end)"`
}

// 行削除リクエスト
type DeleteRowsRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	Count           int64  `json:"count" jsonschema:"required,description=number of rows to delete"`
	StartRow        int64  `json:"start_row" jsonschema:"required,description=row index to start deleting (0-based)"`
}

// 列削除リクエスト
type DeleteColumnsRequest struct {
	SpreadsheetName string `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string `json:"sheet" jsonschema:"required,description=sheet name"`
	Count           int64  `json:"count" jsonschema:"required,description=number of columns to delete"`
	StartColumn     int64  `json:"start_column" jsonschema:"required,description=column index to start deleting (0-based)"`
}

// セル編集リクエスト
type UpdateCellsRequest struct {
	SpreadsheetName string          `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string          `json:"sheet" jsonschema:"required,description=sheet name"`
	Range           string          `json:"range" jsonschema:"required,description=cell range (e.g. A1:C10)"`
	Data            [][]interface{} `json:"data" jsonschema:"required,description=2D array of cell values to update"`
}

// 複数範囲のセル編集リクエスト
type BatchUpdateCellsRequest struct {
	SpreadsheetName string                     `json:"spreadsheet" jsonschema:"required,description=spreadsheet name"`
	SheetName       string                     `json:"sheet" jsonschema:"required,description=sheet name"`
	Ranges          map[string][][]interface{} `json:"ranges" jsonschema:"required,description=map of range to 2D array of cell values (e.g. {'A1:B2': [[1, 2], [3, 4]], 'D5:E6': [[5, 6], [7, 8]]})"`
}

// スプレッドシート名からスプレッドシートIDを取得する
func (gs *GoogleSheets) getSpreadsheetId(spreadsheetName string) (string, error) {
	return gs.getSpreadsheetIdWithContext(context.Background(), spreadsheetName)
}

// コンテキスト付きでスプレッドシート名からスプレッドシートIDを取得する
func (gs *GoogleSheets) getSpreadsheetIdWithContext(ctx context.Context, spreadsheetName string) (string, error) {
	// スプレッドシート名からIDを検索
	query := fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false", gs.cfg.FolderID, spreadsheetName)
	service, err := gs.auth.GetDriveService(ctx)
	if err != nil {
		return "", fmt.Errorf("failed to get drive service: %w", err)
	}

	// 明示的に型を指定
	driveService := service

	fileList, err := driveService.Files.List().
		Q(query).
		SupportsAllDrives(true).
		IncludeItemsFromAllDrives(true).
		Fields("files(id, name)").
		Do()
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

func (gs *GoogleSheets) CopySheetHandler(request CopySheetRequest) (*mcp.ToolResponse, error) {
	return gs.CopySheetHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) CopySheetHandlerWithContext(ctx context.Context, request CopySheetRequest) (*mcp.ToolResponse, error) {
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

	return mcp.NewToolResponse(
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully copied to sheet '%s' in spreadsheet '%s'",
			request.SrcSheetName, request.SrcSpreadsheetName,
			request.DstSheetName, request.DstSpreadsheetName)),
	), nil
}

func (gs *GoogleSheets) RenameSheetHandler(request RenameSheetRequest) (*mcp.ToolResponse, error) {
	return gs.RenameSheetHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) RenameSheetHandlerWithContext(ctx context.Context, request RenameSheetRequest) (*mcp.ToolResponse, error) {
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

	return mcp.NewToolResponse(
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully renamed to '%s'",
			request.SheetName, request.SpreadsheetName, request.NewName)),
	), nil
}

func (gs *GoogleSheets) ListSheetsHandler(request ListSheetsRequest) (*mcp.ToolResponse, error) {
	return gs.ListSheetsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) ListSheetsHandlerWithContext(ctx context.Context, request ListSheetsRequest) (*mcp.ToolResponse, error) {
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
	return mcp.NewToolResponse(
		mcp.NewTextContent(result.String()),
	), nil
}

func (gs *GoogleSheets) AddRowsHandler(request AddRowsRequest) (*mcp.ToolResponse, error) {
	return gs.AddRowsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) AddRowsHandlerWithContext(ctx context.Context, request AddRowsRequest) (*mcp.ToolResponse, error) {
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

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// リクエストを作成
	var batchRequest *sheets.BatchUpdateSpreadsheetRequest

	// StartRowが指定されている場合は、特定の位置に行を挿入
	if request.StartRow > 0 {
		batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{
				{
					InsertDimension: &sheets.InsertDimensionRequest{
						Range: &sheets.DimensionRange{
							SheetId:    sheetId,
							Dimension:  "ROWS",
							StartIndex: request.StartRow,
							EndIndex:   request.StartRow + request.Count,
						},
						InheritFromBefore: false,
					},
				},
			},
		}
	} else {
		// 指定がない場合は末尾に追加
		batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{
				{
					AppendDimension: &sheets.AppendDimensionRequest{
						SheetId:   sheetId,
						Dimension: "ROWS",
						Length:    request.Count,
					},
				},
			},
		}
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

	return mcp.NewToolResponse(
		mcp.NewTextContent(message),
	), nil
}

func (gs *GoogleSheets) AddColumnsHandler(request AddColumnsRequest) (*mcp.ToolResponse, error) {
	return gs.AddColumnsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) AddColumnsHandlerWithContext(ctx context.Context, request AddColumnsRequest) (*mcp.ToolResponse, error) {
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

	// シートIDを取得
	sheetId, err := gs.getSheetIdWithContext(ctx, spreadsheetId, sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get sheet ID: %w", err)
	}

	// リクエストを作成
	var batchRequest *sheets.BatchUpdateSpreadsheetRequest

	// StartColumnが指定されている場合は、特定の位置に列を挿入
	if request.StartColumn > 0 {
		batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{
				{
					InsertDimension: &sheets.InsertDimensionRequest{
						Range: &sheets.DimensionRange{
							SheetId:    sheetId,
							Dimension:  "COLUMNS",
							StartIndex: request.StartColumn,
							EndIndex:   request.StartColumn + request.Count,
						},
						InheritFromBefore: false,
					},
				},
			},
		}
	} else {
		// 指定がない場合は末尾に追加
		batchRequest = &sheets.BatchUpdateSpreadsheetRequest{
			Requests: []*sheets.Request{
				{
					AppendDimension: &sheets.AppendDimensionRequest{
						SheetId:   sheetId,
						Dimension: "COLUMNS",
						Length:    request.Count,
					},
				},
			},
		}
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

	return mcp.NewToolResponse(
		mcp.NewTextContent(message),
	), nil
}

// 2次元配列のデータを表形式の文字列に変換する関数
func formatTableData(rangeStr string, values [][]interface{}) string {
	var builder strings.Builder
	builder.WriteString(fmt.Sprintf("\nRange: %s\n", rangeStr))

	for i, row := range values {
		builder.WriteString(fmt.Sprintf("Row %d: ", i+1))
		for j, cell := range row {
			if j > 0 {
				builder.WriteString(" | ")
			}
			builder.WriteString(fmt.Sprintf("%v", cell))
		}
		builder.WriteString("\n")
	}

	return builder.String()
}

// 列インデックス（0-based）をA1表記の列文字（A, B, C, ...）に変換する関数
func columnIndexToLetter(index int64) string {
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
func (gs *GoogleSheets) UpdateCellsHandler(request UpdateCellsRequest) (*mcp.ToolResponse, error) {
	return gs.UpdateCellsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) UpdateCellsHandlerWithContext(ctx context.Context, request UpdateCellsRequest) (*mcp.ToolResponse, error) {
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
	if prevRowCount > 0 {
		prevColCount = len(prevData.Values[0])
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
	prevDataStr := "\n\nPrevious data details:" + formatTableData(request.Range, prevData.Values)

	// レスポンスを作成（変更前のデータを含める）
	return mcp.NewToolResponse(
		mcp.NewTextContent(message + prevDataStr),
	), nil
}

// 複数範囲のセル一括編集ハンドラー
func (gs *GoogleSheets) BatchUpdateCellsHandler(request BatchUpdateCellsRequest) (*mcp.ToolResponse, error) {
	return gs.BatchUpdateCellsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) BatchUpdateCellsHandlerWithContext(ctx context.Context, request BatchUpdateCellsRequest) (*mcp.ToolResponse, error) {
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
	prevDataStr.WriteString("\n\nPrevious data details:")
	for rangeStr, values := range previousData {
		prevDataStr.WriteString(formatTableData(rangeStr, values))
	}

	// レスポンスを作成（変更前のデータを含める）
	return mcp.NewToolResponse(
		mcp.NewTextContent(message + prevDataStr.String()),
	), nil
}

func (gs *GoogleSheets) GetSheetDataHandler(request GetSheetDataRequest) (*mcp.ToolResponse, error) {
	return gs.GetSheetDataHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) GetSheetDataHandlerWithContext(ctx context.Context, request GetSheetDataRequest) (*mcp.ToolResponse, error) {
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

// 行削除ハンドラー
func (gs *GoogleSheets) DeleteRowsHandler(request DeleteRowsRequest) (*mcp.ToolResponse, error) {
	return gs.DeleteRowsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) DeleteRowsHandlerWithContext(ctx context.Context, request DeleteRowsRequest) (*mcp.ToolResponse, error) {
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
	if request.StartRow < 0 {
		return nil, fmt.Errorf("start row must be a non-negative number")
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
	startRowA1 := request.StartRow + 1 // 0-based to 1-based
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
						StartIndex: request.StartRow,
						EndIndex:   request.StartRow + request.Count,
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
	prevDataStr := "\n\nDeleted data details:" + formatTableData(fmt.Sprintf("%d:%d", startRowA1, endRowA1), prevData.Values)

	return mcp.NewToolResponse(
		mcp.NewTextContent(message + prevDataStr),
	), nil
}

// 列削除ハンドラー
func (gs *GoogleSheets) DeleteColumnsHandler(request DeleteColumnsRequest) (*mcp.ToolResponse, error) {
	return gs.DeleteColumnsHandlerWithContext(context.Background(), request)
}

func (gs *GoogleSheets) DeleteColumnsHandlerWithContext(ctx context.Context, request DeleteColumnsRequest) (*mcp.ToolResponse, error) {
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
	if request.StartColumn < 0 {
		return nil, fmt.Errorf("start column must be a non-negative number")
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
						StartIndex: request.StartColumn,
						EndIndex:   request.StartColumn + request.Count,
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
	if prevRowCount > 0 {
		prevColCount = len(prevData.Values[0])
	}

	message += fmt.Sprintf("\n\nDeleted data (%d rows x %d columns) has been saved. To undo this change, you can use the saved data.",
		prevRowCount, prevColCount)

	// 削除前のデータを表示用に整形
	prevDataStr := "\n\nDeleted data details:" + formatTableData(fmt.Sprintf("%s:%s", startColA1, endColA1), prevData.Values)

	return mcp.NewToolResponse(
		mcp.NewTextContent(message + prevDataStr),
	), nil
}
