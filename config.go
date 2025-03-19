package main

import (
	"fmt"

	"github.com/kelseyhightower/envconfig"
)

const envPrefix = "GSUITE"

type Config struct {
	ClientSecretPath string `envconfig:"CLIENT_SECRET_PATH"`
	TokenPath        string `envconfig:"TOKEN_PATH"`
	FolderID         string `envconfig:"FOLDER_ID"`
}

func NewConfig() (*Config, error) {
	c := &Config{}
	err := envconfig.Process(envPrefix, c)
	if err != nil {
		return nil, fmt.Errorf("failed to parse environment variables: %w", err)
	}
	return c, nil
}
