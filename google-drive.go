package main

import (
	"context"
	"fmt"
	"net/http"

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

type CopyFileInput struct {
	SrcPath string `json:"src_path" jsonschema:"required,description=source path"`
	DstPath string `json:"dst_path" jsonschema:"required,description=destination path"`
}

type CopyFileOutput struct{}

func (gd *GoogleDrive) CopyFileHandler(input CopyFileInput) (CopyFileOutput, error) {
	// TODO: Implement
	return CopyFileOutput{}, nil
}
