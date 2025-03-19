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
	// スプレッドシート名からIDを検索
	query := fmt.Sprintf("'%s' in parents and name = '%s' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false", gs.cfg.FolderID, spreadsheetName)
	fileList, err := gs.drive.Files.List().
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
	// ソーススプレッドシートのIDを取得
	srcSpreadsheetId, err := gs.getSpreadsheetId(request.SrcSpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get source spreadsheet ID: %w", err)
	}

	// 宛先スプレッドシートのIDを取得
	dstSpreadsheetId, err := gs.getSpreadsheetId(request.DstSpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get destination spreadsheet ID: %w", err)
	}

	srcSheetName := request.SrcSheetName
	dstSheetName := request.DstSheetName

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
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully copied to sheet '%s' in spreadsheet '%s'",
			request.SrcSheetName, request.SrcSpreadsheetName,
			request.DstSheetName, request.DstSpreadsheetName)),
	), nil
}

func (gs *GoogleSheets) RenameSheetHandler(request RenameSheetRequest) (*mcp.ToolResponse, error) {
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

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
		mcp.NewTextContent(fmt.Sprintf("Sheet '%s' in spreadsheet '%s' successfully renamed to '%s'",
			request.SheetName, request.SpreadsheetName, request.NewName)),
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

func (gs *GoogleSheets) AddRowsHandler(request AddRowsRequest) (*mcp.ToolResponse, error) {
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 追加する行数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetId(spreadsheetId, sheetName)
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
	_, err = gs.service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
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
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 追加する列数が正の値であることを確認
	if request.Count <= 0 {
		return nil, fmt.Errorf("count must be a positive number")
	}

	// シートIDを取得
	sheetId, err := gs.getSheetId(spreadsheetId, sheetName)
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
	_, err = gs.service.Spreadsheets.BatchUpdate(spreadsheetId, batchRequest).Do()
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

// 単一範囲のセル編集ハンドラー
func (gs *GoogleSheets) UpdateCellsHandler(request UpdateCellsRequest) (*mcp.ToolResponse, error) {
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
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

	// 値を更新するリクエストを作成
	valueRange := &sheets.ValueRange{
		Range:  fullRange,
		Values: request.Data,
	}

	// 値を更新
	updateResponse, err := gs.service.Spreadsheets.Values.Update(
		spreadsheetId, fullRange, valueRange).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		return nil, fmt.Errorf("failed to update cells: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully updated %d cells in range '%s' of sheet '%s' in spreadsheet '%s'",
		updateResponse.UpdatedCells, request.Range, request.SheetName, request.SpreadsheetName)

	return mcp.NewToolResponse(
		mcp.NewTextContent(message),
	), nil
}

// 複数範囲のセル一括編集ハンドラー
func (gs *GoogleSheets) BatchUpdateCellsHandler(request BatchUpdateCellsRequest) (*mcp.ToolResponse, error) {
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet ID: %w", err)
	}

	sheetName := request.SheetName

	// 範囲が指定されていることを確認
	if len(request.Ranges) == 0 {
		return nil, fmt.Errorf("ranges cannot be empty")
	}

	// バッチ更新用のデータを作成
	var data []*sheets.ValueRange
	for rangeStr, values := range request.Ranges {
		// データが空でないことを確認
		if len(values) == 0 {
			return nil, fmt.Errorf("data for range '%s' cannot be empty", rangeStr)
		}

		// 範囲を完全な形式に変換（シート名を含む）
		fullRange := fmt.Sprintf("%s!%s", sheetName, rangeStr)

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
	batchUpdateResponse, err := gs.service.Spreadsheets.Values.BatchUpdate(
		spreadsheetId, batchUpdateRequest).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to batch update cells: %w", err)
	}

	// 成功メッセージを作成
	message := fmt.Sprintf("Successfully updated %d cells across %d ranges in sheet '%s' in spreadsheet '%s'",
		batchUpdateResponse.TotalUpdatedCells, batchUpdateResponse.TotalUpdatedSheets,
		request.SheetName, request.SpreadsheetName)

	return mcp.NewToolResponse(
		mcp.NewTextContent(message),
	), nil
}

func (gs *GoogleSheets) GetSheetDataHandler(request GetSheetDataRequest) (*mcp.ToolResponse, error) {
	// スプレッドシートIDを取得
	spreadsheetId, err := gs.getSpreadsheetId(request.SpreadsheetName)
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
	resp, err := gs.service.Spreadsheets.Values.Get(spreadsheetId, range_).Do()
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
