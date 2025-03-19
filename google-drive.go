package main

import (
	"context"
	"fmt"
	"net/http"

	mcp "github.com/metoro-io/mcp-golang"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
)

type GoogleDrive struct {
	cfg     *Config
	service *drive.Service
}

func NewGoogleDrive(ctx context.Context, cfg *Config, client *http.Client) (*GoogleDrive, error) {
	service, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		return nil, fmt.Errorf("failed to create drive service: %w", err)
	}
	return &GoogleDrive{
		cfg:     cfg,
		service: service,
	}, nil
}

type CopyFileRequest struct {
	SrcPath string `json:"src_path" jsonschema:"required,description=source path"`
	DstPath string `json:"dst_path" jsonschema:"required,description=destination path"`
}

func (gd *GoogleDrive) CopyFileHandler(request CopyFileRequest) (*mcp.ToolResponse, error) {
	// TODO: Implement
	return mcp.NewToolResponse(
		mcp.NewTextContent("success"),
	), nil
}
