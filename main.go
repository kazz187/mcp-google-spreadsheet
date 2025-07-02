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
	server := mcp.NewServer("mcp-google-spreadsheet", "v1.0.0", nil)

	// Register all tools
	server.AddTools(
		mcp.NewServerTool(
			"list_files",
			"List files in google drive",
			drive.ListFilesHandler,
			ListFilesInputSchema,
		),
		mcp.NewServerTool(
			"copy_file",
			"Copy file in google drive",
			drive.CopyFileHandler,
			CopyFileInputSchema,
		),
		mcp.NewServerTool(
			"rename_file",
			"Rename file in google drive",
			drive.RenameFileHandler,
			RenameFileInputSchema,
		),
		mcp.NewServerTool(
			"list_sheets",
			"List sheets in google sheet",
			sheet.ListSheetsHandler,
			ListSheetsInputSchema,
		),
		mcp.NewServerTool(
			"copy_sheet",
			"Copy sheet in google sheet",
			sheet.CopySheetHandler,
			CopySheetInputSchema,
		),
		mcp.NewServerTool(
			"rename_sheet",
			"Rename sheet in google sheet",
			sheet.RenameSheetHandler,
			RenameSheetInputSchema,
		),
		mcp.NewServerTool(
			"get_sheet_data",
			"Get data from sheet in google sheet",
			sheet.GetSheetDataHandler,
			GetSheetDataInputSchema,
		),
		mcp.NewServerTool(
			"add_rows",
			"Add rows to sheet in google sheet",
			sheet.AddRowsHandler,
			AddRowsInputSchema,
		),
		mcp.NewServerTool(
			"add_columns",
			"Add columns to sheet in google sheet",
			sheet.AddColumnsHandler,
			AddColumnsInputSchema,
		),
		mcp.NewServerTool(
			"update_cells",
			"Update cells in google sheet",
			sheet.UpdateCellsHandler,
			UpdateCellsInputSchema,
		),
		mcp.NewServerTool(
			"batch_update_cells",
			"Batch update cells in google sheet",
			sheet.BatchUpdateCellsHandler,
			BatchUpdateCellsInputSchema,
		),
		mcp.NewServerTool(
			"delete_rows",
			"Delete rows from sheet in google sheet",
			sheet.DeleteRowsHandler,
			DeleteRowsInputSchema,
		),
		mcp.NewServerTool(
			"delete_columns",
			"Delete columns from sheet in google sheet",
			sheet.DeleteColumnsHandler,
			DeleteColumnsInputSchema,
		),
	)

	if err := server.Run(ctx, mcp.NewStdioTransport()); err != nil {
		logger.ErrorContext(ctx, "failed to run server", "error", err)
		os.Exit(1)
	}
}
