package main

import (
	"context"
	"log/slog"
	"os"
	"os/signal"
	"syscall"

	mcp "github.com/metoro-io/mcp-golang"
	"github.com/metoro-io/mcp-golang/transport/stdio"
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
	server := mcp.NewServer(stdio.NewStdioServerTransport())
	if err := server.RegisterTool("list_files", "List files in google drive", drive.ListFilesHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool list_files", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("copy_file", "Copy file in google drive", drive.CopyFileHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool copy_file", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("copy_sheet", "Copy sheet in google sheet", sheet.CopySheetHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool copy_sheet", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("rename_file", "Rename file in google drive", drive.RenameFileHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool rename_file", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("rename_sheet", "Rename sheet in google sheet", sheet.RenameSheetHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool rename_sheet", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("list_sheets", "List sheets in google sheet", sheet.ListSheetsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool list_sheets", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("get_sheet_data", "Get data from sheet in google sheet", sheet.GetSheetDataHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool get_sheet_data", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("add_rows", "Add rows to sheet in google sheet", sheet.AddRowsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool add_rows", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("add_columns", "Add columns to sheet in google sheet", sheet.AddColumnsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool add_columns", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("update_cells", "Update cells in google sheet", sheet.UpdateCellsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool update_cells", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("batch_update_cells", "Batch update cells in google sheet", sheet.BatchUpdateCellsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool batch_update_cells", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("delete_rows", "Delete rows from sheet in google sheet", sheet.DeleteRowsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool delete_rows", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("delete_columns", "Delete columns from sheet in google sheet", sheet.DeleteColumnsHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool delete_columns", "error", err)
		os.Exit(1)
	}
	if err := server.Serve(); err != nil {
		logger.ErrorContext(ctx, "failed to serve", "error", err)
		os.Exit(1)
	}
	<-ctx.Done()
}
