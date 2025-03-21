package main

import (
	"fmt"
	"os"

	"github.com/kelseyhightower/envconfig"
)

const envPrefix = "MCPGS"

type Config struct {
	ClientSecretPath string `envconfig:"CLIENT_SECRET_PATH"`
	TokenPathRaw     string `envconfig:"TOKEN_PATH"`
	TokenPath        string `envconfig:"-"`
	FolderID         string `envconfig:"FOLDER_ID"`
}

func NewConfig() (*Config, error) {
	c := &Config{}
	err := envconfig.Process(envPrefix, c)
	if err != nil {
		return nil, fmt.Errorf("failed to parse environment variables: %w", err)
	}
	// ローカルにトークンが保存されていれば、それを使う
	tokenPath := c.TokenPathRaw
	if tokenPath == "" {
		homeDir, err := os.UserHomeDir()
		if err != nil {
			return nil, fmt.Errorf("failed to determine home directory: %w", err)
		}
		tokenPath = homeDir + "/.mcp_google_spreadsheet.json"
	}
	return c, nil
}
