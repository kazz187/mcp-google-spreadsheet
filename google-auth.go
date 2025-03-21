package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"os"
	"os/exec"
	"runtime"
	"time"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	drive "google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type GoogleAuth struct {
	cfg    *Config
	config *oauth2.Config
	token  *oauth2.Token
}

func NewGoogleAuth(cfg *Config) *GoogleAuth {
	return &GoogleAuth{
		cfg: cfg,
	}
}

func (g *GoogleAuth) AuthClient(ctx context.Context) (*http.Client, error) {
	b, err := os.ReadFile(g.cfg.ClientSecretPath)
	if err != nil {
		return nil, fmt.Errorf("could not read client id file: %w", err)
	}
	gCfg, err := google.ConfigFromJSON(b, sheets.DriveScope, sheets.SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("unable to parse client id file to config: %w", err)
	}

	// OAuth2設定を保存
	g.config = gCfg

	client, err := g.getClient(ctx, gCfg)
	if err != nil {
		return nil, fmt.Errorf("failed to get client: %w", err)
	}
	return client, nil
}

// GetSheetsService は認証済みのSheetsサービスを返します
func (g *GoogleAuth) GetSheetsService(ctx context.Context) (*sheets.Service, error) {
	client, err := g.getClient(ctx, g.config)
	if err != nil {
		return nil, fmt.Errorf("failed to get client: %w", err)
	}

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		newClient, refreshErr := g.refreshAndGetClient(ctx)
		if refreshErr != nil {
			return nil, fmt.Errorf("failed to refresh token after API error: %w, %w", err, refreshErr)
		}
		return sheets.NewService(ctx, option.WithHTTPClient(newClient))
	}
	return srv, nil
}

// GetDriveService は認証済みのDriveサービスを返します
func (g *GoogleAuth) GetDriveService(ctx context.Context) (*drive.Service, error) {
	client, err := g.getClient(ctx, g.config)
	if err != nil {
		return nil, fmt.Errorf("failed to get client: %w", err)
	}

	srv, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		newClient, refreshErr := g.refreshAndGetClient(ctx)
		if refreshErr != nil {
			return nil, fmt.Errorf("failed to refresh token after API error: %w, %w", err, refreshErr)
		}
		return drive.NewService(ctx, option.WithHTTPClient(newClient))
	}
	return srv, nil
}

// refreshAndGetClient はトークンをリフレッシュして新しいクライアントを返します
func (g *GoogleAuth) refreshAndGetClient(ctx context.Context) (*http.Client, error) {
	if g.config == nil {
		// configがまだ初期化されていない場合は初期化
		b, err := os.ReadFile(g.cfg.ClientSecretPath)
		if err != nil {
			return nil, fmt.Errorf("could not read client id file: %w", err)
		}
		gCfg, err := google.ConfigFromJSON(b, sheets.DriveScope, sheets.SpreadsheetsScope)
		if err != nil {
			return nil, fmt.Errorf("unable to parse client id file to config: %w", err)
		}
		g.config = gCfg
	}

	// 再認証を試みる
	fmt.Println("Attempting to re-authenticate...")
	tok, err := getTokenFromWeb(ctx, g.config)
	if err != nil {
		return nil, fmt.Errorf("unable to get new token: %w", err)
	}

	if err := saveToken(g.cfg.TokenPath, tok); err != nil {
		return nil, fmt.Errorf("unable to save token: %w", err)
	}

	return g.config.Client(ctx, tok), nil
}

// 認証情報を取得し、HTTP クライアントを作成
func (g *GoogleAuth) getClient(ctx context.Context, config *oauth2.Config) (*http.Client, error) {
	if g.token == nil {
		tok, err := tokenFromFile(g.cfg.TokenPath)
		if err != nil {
			// トークンがない場合、新しく取得する
			tok, err := getTokenFromWeb(ctx, config)
			if err != nil {
				return nil, fmt.Errorf("unable to get token: %w", err)
			}
			if err := saveToken(g.cfg.TokenPath, tok); err != nil {
				return nil, fmt.Errorf("unable to save token: %w", err)
			}
		}
		g.token = tok
	}

	// トークンの有効期限をチェック
	if g.token.Expiry.Before(time.Now()) {
		fmt.Println("Token has expired, refreshing...")

		// TokenSourceを使ってトークンをリフレッシュ
		tokenSource := config.TokenSource(ctx, g.token)
		newToken, err := tokenSource.Token()
		if err != nil {
			fmt.Printf("Failed to refresh token: %v\n", err)
			// リフレッシュに失敗した場合は再認証
			newToken, err = getTokenFromWeb(ctx, config)
			if err != nil {
				return nil, fmt.Errorf("unable to get new token: %w", err)
			}
		}
		// 新しいトークンを保存
		if err := saveToken(g.cfg.TokenPath, newToken); err != nil {
			return nil, fmt.Errorf("unable to save refreshed token: %w", err)
		}
		g.token = newToken
	}

	return config.Client(ctx, g.token), nil
}

// ブラウザで認証し、認証コードを取得
func getTokenFromWeb(ctx context.Context, config *oauth2.Config) (*oauth2.Token, error) {
	// localhost でリダイレクトを受け取るように設定
	config.RedirectURL = "http://localhost:8080/oauth2callback"

	// 認証コードを受け取るためのチャネル
	codeCh := make(chan string, 1)
	errCh := make(chan error, 1)

	// 一時的なHTTPサーバーを起動
	server := &http.Server{Addr: ":8080"}

	// コールバックハンドラー
	http.HandleFunc("/oauth2callback", func(w http.ResponseWriter, r *http.Request) {
		code := r.URL.Query().Get("code")
		if code == "" {
			errCh <- fmt.Errorf("no code in request")
			http.Error(w, "No code in request", http.StatusBadRequest)
			return
		}

		// 認証成功メッセージを表示
		w.Header().Set("Content-Type", "text/html")
		w.Write([]byte(`
			<html>
				<head>
					<title>Authentication Successful</title>
					<style>
						body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
						.success { color: #4CAF50; font-size: 24px; margin-bottom: 20px; }
						.message { font-size: 16px; margin-bottom: 30px; }
					</style>
				</head>
				<body>
					<div class="success">Authentication Successful!</div>
					<div class="message">You can close this window and return to the application.</div>
				</body>
			</html>
		`))

		// コードをチャネルに送信
		codeCh <- code

		// 5秒後にサーバーをシャットダウン
		go func() {
			time.Sleep(1 * time.Second)
			server.Shutdown(context.Background())
		}()
	})

	// サーバーを別のゴルーチンで起動
	go func() {
		if err := server.ListenAndServe(); !errors.Is(err, http.ErrServerClosed) {
			errCh <- err
		}
	}()

	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)

	// ブラウザを自動的に開く
	fmt.Printf("Opening browser for authentication: %s\n", authURL)
	if err := openBrowser(authURL); err != nil {
		fmt.Printf("Could not open browser automatically. Please open the following URL in your browser: \n%s\n", authURL)
	}

	// コードまたはエラーを待つ
	select {
	case code := <-codeCh:
		// 認証コードを使ってアクセストークンを取得
		tok, err := config.Exchange(ctx, code)
		if err != nil {
			return nil, fmt.Errorf("failed to exchange authorization code: %w", err)
		}
		return tok, nil
	case err := <-errCh:
		return nil, fmt.Errorf("error during authentication process: %w", err)
	case <-time.After(2 * time.Minute):
		return nil, fmt.Errorf("authentication timed out")
	}
}

// ブラウザを開く関数
func openBrowser(url string) error {
	var cmd *exec.Cmd

	switch runtime.GOOS {
	case "darwin":
		cmd = exec.Command("open", url)
	case "windows":
		cmd = exec.Command("rundll32", "url.dll,FileProtocolHandler", url)
	case "linux":
		cmd = exec.Command("xdg-open", url)
	default:
		return fmt.Errorf("unsupported platform")
	}

	return cmd.Start()
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
		return fmt.Errorf("failed creating token file: %w", err)
	}
	defer f.Close()
	if err := json.NewEncoder(f).Encode(token); err != nil {
		return fmt.Errorf("failed encoding token: %w", err)
	}
	return nil
}
