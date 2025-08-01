package main

import (
	"context"
	"log/slog"
	"os"
	"os/signal"
	"syscall"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func main() {
	logger := slog.Default()
	ctx, cancel := signal.NotifyContext(context.Background(), os.Interrupt, syscall.SIGTERM)
	defer cancel()
	cfg, err := NewConfig()
	if err != nil {
		logger.ErrorContext(ctx, "failed to create config", "error", err)
		os.Exit(1)
	}
	gAuth := NewGoogleAuth(cfg)
	// 認証クライアントを取得（GoogleAuthの初期化のため）
	_, err = gAuth.AuthClient(ctx)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create auth client", "error", err)
		os.Exit(1)
	}
	drive, err := NewGoogleDrive(cfg, gAuth)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create drive", "error", err)
		os.Exit(1)
	}
	sheet, err := NewGoogleSheets(cfg, gAuth)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create sheet", "error", err)
		os.Exit(1)
	}
	server := mcp.NewServer(
		&mcp.Implementation{
			Name:    "mcp-google-spreadsheet",
			Title:   "Google Drive & Sheets MCP Server",
			Version: "v1.0.0",
		},
		&mcp.ServerOptions{
			Instructions: "MCP server for Google Drive file management and Google Sheets operations. Workflow: 1) Use google_drive_list_files to browse and find spreadsheets, 2) Use google_sheets_list_sheets to see sheets in a spreadsheet, 3) Use google_sheets_read_data to view content, 4) Use other google_sheets_* tools to modify data.",
		},
	)

	// Register Google Drive tools
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_drive_list_files",
			Title:       "Google Drive: List Files and Folders",
			Description: "Browse and list files and folders in Google Drive. Use this to explore directory structure and find spreadsheets before working with them.",
			InputSchema: ListFilesInputSchema,
		},
		drive.ListFilesHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_drive_copy_file",
			Title:       "Google Drive: Copy File",
			Description: "Copy a file or folder to another location in Google Drive. Specify source and destination paths.",
			InputSchema: CopyFileInputSchema,
		},
		drive.CopyFileHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_drive_rename_file",
			Title:       "Google Drive: Rename File",
			Description: "Rename a file or folder in Google Drive. Provide the current file path and new name.",
			InputSchema: RenameFileInputSchema,
		},
		drive.RenameFileHandler,
	)
	// Register Google Sheets tools
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_list_sheets",
			Title:       "Google Sheets: List Sheets in Spreadsheet",
			Description: "List all sheets (tabs) within a specific Google Spreadsheet. Use this after finding the spreadsheet with google_drive_list_files.",
			InputSchema: ListSheetsInputSchema,
		},
		sheet.ListSheetsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_copy_sheet",
			Title:       "Google Sheets: Copy Sheet",
			Description: "Copy a sheet from one Google Spreadsheet to another. Specify source and destination spreadsheet names and sheet names.",
			InputSchema: CopySheetInputSchema,
		},
		sheet.CopySheetHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_rename_sheet",
			Title:       "Google Sheets: Rename Sheet",
			Description: "Rename a sheet (tab) within a Google Spreadsheet. Provide spreadsheet name, current sheet name, and new name.",
			InputSchema: RenameSheetInputSchema,
		},
		sheet.RenameSheetHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_read_data",
			Title:       "Google Sheets: Read Data from Sheet",
			Description: "Read data from a specific sheet in a Google Spreadsheet. Specify spreadsheet name, sheet name, and optionally a cell range (e.g., A1:C10). This is how you 'open' and view spreadsheet content.",
			InputSchema: GetSheetDataInputSchema,
		},
		sheet.GetSheetDataHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_add_rows",
			Title:       "Google Sheets: Insert Rows",
			Description: "Insert new empty rows in a Google Sheet. Specify spreadsheet name, sheet name, number of rows to add, and starting row position.",
			InputSchema: AddRowsInputSchema,
		},
		sheet.AddRowsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_add_columns",
			Title:       "Google Sheets: Insert Columns",
			Description: "Insert new empty columns in a Google Sheet. Specify spreadsheet name, sheet name, number of columns to add, and starting column position.",
			InputSchema: AddColumnsInputSchema,
		},
		sheet.AddColumnsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_update_cells",
			Title:       "Google Sheets: Update Cell Values",
			Description: "Update cell values in a specific range of a Google Sheet. Provide spreadsheet name, sheet name, cell range (e.g., A1:C3), and 2D array of values.",
			InputSchema: UpdateCellsInputSchema,
		},
		sheet.UpdateCellsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_batch_update_cells",
			Title:       "Google Sheets: Batch Update Multiple Ranges",
			Description: "Update multiple cell ranges in a Google Sheet in a single operation. Provide spreadsheet name, sheet name, and a map of ranges to values.",
			InputSchema: BatchUpdateCellsInputSchema,
		},
		sheet.BatchUpdateCellsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_delete_rows",
			Title:       "Google Sheets: Delete Rows",
			Description: "Delete rows from a Google Sheet. Specify spreadsheet name, sheet name, number of rows to delete, and starting row position.",
			InputSchema: DeleteRowsInputSchema,
		},
		sheet.DeleteRowsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "google_sheets_delete_columns",
			Title:       "Google Sheets: Delete Columns",
			Description: "Delete columns from a Google Sheet. Specify spreadsheet name, sheet name, number of columns to delete, and starting column position.",
			InputSchema: DeleteColumnsInputSchema,
		},
		sheet.DeleteColumnsHandler,
	)

	if err := server.Run(ctx, mcp.NewStdioTransport()); err != nil {
		logger.ErrorContext(ctx, "failed to run server", "error", err)
		os.Exit(1)
	}
}
