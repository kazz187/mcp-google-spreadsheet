package main

import (
	"context"
	"fmt"
	"net/http"

	mcp "github.com/metoro-io/mcp-golang"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type GoogleSheets struct {
	cfg     *Config
	service *sheets.Service
}

func NewGoogleSheets(ctx context.Context, cfg *Config, cli *http.Client) (*GoogleSheets, error) {
	service, err := sheets.NewService(ctx, option.WithHTTPClient(cli))
	if err != nil {
		return nil, fmt.Errorf("failed to create sheets service: %w", err)
	}
	return &GoogleSheets{
		cfg:     cfg,
		service: service,
	}, nil
}

type CopySheetRequest struct {
	SrcName string `json:"src_path" jsonschema:"required,description=source sheet name"`
	DstName string `json:"dst_path" jsonschema:"required,description=destination sheet name"`
}

func (gd *GoogleSheets) CopySheetHandler(request CopySheetRequest) (*mcp.ToolResponse, error) {
	// TODO: Implement
	return mcp.NewToolResponse(
		mcp.NewTextContent("success"),
	), nil
}
