package main

import (
	"context"
	"fmt"
	"net/http"

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

type CopySheetInput struct {
	SrcName string `json:"src_path" jsonschema:"required,description=source sheet name"`
	DstName string `json:"dst_path" jsonschema:"required,description=destination sheet name"`
}

type CopySheetOutput struct{}

func (gd *GoogleSheets) CopySheetHandler(input CopySheetInput) (CopySheetOutput, error) {
	return CopySheetOutput{}, nil
}
