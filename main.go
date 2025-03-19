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
	authCli, err := gAuth.AuthClient(ctx)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create auth client", "error", err)
		os.Exit(1)
	}
	drive, err := NewGoogleDrive(ctx, cfg, authCli)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create drive", "error", err)
		os.Exit(1)
	}
	sheet, err := NewGoogleSheets(ctx, cfg, authCli)
	if err != nil {
		logger.ErrorContext(ctx, "failed to create sheet", "error", err)
		os.Exit(1)
	}
	server := mcp.NewServer(stdio.NewStdioServerTransport())
	if err := server.RegisterTool("copy_file", "Copy file in google drive", drive.CopyFileHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool copy_file", "error", err)
		os.Exit(1)
	}
	if err := server.RegisterTool("copy_sheet", "Copy sheet in google sheet", sheet.CopySheetHandler); err != nil {
		logger.ErrorContext(ctx, "failed to register tool copy_sheet", "error", err)
		os.Exit(1)
	}
	if err := server.Serve(); err != nil {
		logger.ErrorContext(ctx, "failed to serve", "error", err)
		os.Exit(1)
	}
	<-ctx.Done()
}
