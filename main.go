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
			Title:   "Google Spreadsheet MCP Server",
			Version: "v1.0.0",
		},
		&mcp.ServerOptions{
			Instructions: "MCP server for Google Spreadsheet and Google Drive operations",
		},
	)

	// Register all tools
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "list_files",
			Title:       "List Files",
			Description: "List files in google drive",
			InputSchema: ListFilesInputSchema,
		},
		drive.ListFilesHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "copy_file",
			Title:       "Copy File",
			Description: "Copy file in google drive",
			InputSchema: CopyFileInputSchema,
		},
		drive.CopyFileHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "rename_file",
			Title:       "Rename File",
			Description: "Rename file in google drive",
			InputSchema: RenameFileInputSchema,
		},
		drive.RenameFileHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "list_sheets",
			Title:       "List Sheets",
			Description: "List sheets in google sheet",
			InputSchema: ListSheetsInputSchema,
		},
		sheet.ListSheetsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "copy_sheet",
			Title:       "Copy Sheet",
			Description: "Copy sheet in google sheet",
			InputSchema: CopySheetInputSchema,
		},
		sheet.CopySheetHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "rename_sheet",
			Title:       "Rename Sheet",
			Description: "Rename sheet in google sheet",
			InputSchema: RenameSheetInputSchema,
		},
		sheet.RenameSheetHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "get_sheet_data",
			Title:       "Get Sheet Data",
			Description: "Get data from sheet in google sheet",
			InputSchema: GetSheetDataInputSchema,
		},
		sheet.GetSheetDataHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "add_rows",
			Title:       "Add Rows",
			Description: "Add rows to sheet in google sheet",
			InputSchema: AddRowsInputSchema,
		},
		sheet.AddRowsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "add_columns",
			Title:       "Add Columns",
			Description: "Add columns to sheet in google sheet",
			InputSchema: AddColumnsInputSchema,
		},
		sheet.AddColumnsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "update_cells",
			Title:       "Update Cells",
			Description: "Update cells in google sheet",
			InputSchema: UpdateCellsInputSchema,
		},
		sheet.UpdateCellsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "batch_update_cells",
			Title:       "Batch Update Cells",
			Description: "Batch update cells in google sheet",
			InputSchema: BatchUpdateCellsInputSchema,
		},
		sheet.BatchUpdateCellsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "delete_rows",
			Title:       "Delete Rows",
			Description: "Delete rows from sheet in google sheet",
			InputSchema: DeleteRowsInputSchema,
		},
		sheet.DeleteRowsHandler,
	)
	mcp.AddTool(
		server,
		&mcp.Tool{
			Name:        "delete_columns",
			Title:       "Delete Columns",
			Description: "Delete columns from sheet in google sheet",
			InputSchema: DeleteColumnsInputSchema,
		},
		sheet.DeleteColumnsHandler,
	)

	if err := server.Run(ctx, mcp.NewStdioTransport()); err != nil {
		logger.ErrorContext(ctx, "failed to run server", "error", err)
		os.Exit(1)
	}
}

func ToPtr[T any](v T) *T {
        return &v
}

