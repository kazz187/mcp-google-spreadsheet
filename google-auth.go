package main

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

type GoogleAuth struct {
	cfg *Config
}

func NewGoogleAuth(cfg *Config) *GoogleAuth {
	return &GoogleAuth{
		cfg: cfg,
	}
}

func (g *GoogleAuth) AuthClient(ctx context.Context) (*http.Client, error) {
	b, err := os.ReadFile(g.cfg.ClientSecretPath)
	if err != nil {
		return nil, fmt.Errorf("could not read client secret file: %v", err)
	}
	gCfg, err := google.ConfigFromJSON(b, sheets.DriveScope, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to parse client secret file to config: %v", err)
	}

	client, err := getClient(ctx, gCfg)
	return client, nil
}

// 認証情報を取得し、HTTP クライアントを作成
func getClient(ctx context.Context, config *oauth2.Config) (*http.Client, error) {
	// ローカルにトークンが保存されていれば、それを使う
	homeDir, err := os.UserHomeDir()
	if err != nil {
		return nil, fmt.Errorf("failed to get user home directory: %v", err)
	}
	tokFile := homeDir + "/.mcp_google_spreadsheet.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		// トークンがない場合、新しく取得する
		tok, err := getTokenFromWeb(config)
		if err != nil {
			return nil, fmt.Errorf("unable to get token: %v", err)
		}
		if err := saveToken(tokFile, tok); err != nil {
			return nil, fmt.Errorf("unable to save token: %v", err)
		}
	}
	return config.Client(ctx, tok), nil
}

// ブラウザで認証し、認証コードを取得
func getTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Open the following URL in your browser and then type the auth code: \n%v\n", authURL)

	// ユーザーに認証コードを入力してもらう
	fmt.Print("Input the auth code: ")
	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("failed to read auth code: %v", err)
	}

	// 認証コードを使ってアクセストークンを取得
	tok, err := config.Exchange(context.Background(), authCode)
	if err != nil {
		return nil, fmt.Errorf("failed to exchange auth code: %v", err)
	}
	return tok, nil
}

// ローカルに保存されているトークンを読み込む
func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

// 取得したトークンをローカルファイルに保存
func saveToken(file string, token *oauth2.Token) error {
	f, err := os.Create(file)
	if err != nil {
		return fmt.Errorf("failed creating token file: %v", err)
	}
	defer f.Close()
	if err := json.NewEncoder(f).Encode(token); err != nil {
		return fmt.Errorf("failed encoding token: %v", err)
	}
	return nil
}
