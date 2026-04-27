# 佐大OTP Watcher

佐賀大学SAML認証時に、登録メールに届くワンタイムパスワード（OTP）を自動でクリップボードにコピーする常駐ツール。

## 仕組み

1. ブラウザ拡張機能が `ssoidp.cc.saga-u.ac.jp/idp/profile/SAML2/Redirect` へのアクセスを検知
2. `localhost:18924` で待機する Python スクリプトに HTTP POST を送信
3. Python スクリプトが IMAP でメールサーバーをポーリングし、OTPメールを検索
4. 本文からOTPを抽出してクリップボードにコピー、Toast通知を表示

検知遅延は **0.1秒未満**（拡張機能の `webNavigation` API による即時検知）。

## 必要環境

- Windows 10 以降（Windows 11 推奨）
- Python 3.9 以降（同梱の `python-manager-26.1.msix` で自動インストール可能）
- ブラウザ: Brave / Chrome / Edge / Firefox / Floorp 等
- メールアカウント（IMAP対応）: iCloud / Gmail / Outlook / Yahoo!メール / 任意のIMAPサーバー

## セットアップ手順

### 1. Python のインストール

Python が入っていない場合は、`install_task.bat` を実行すると同梱の `python-manager-26.1.msix` から自動インストールされる。インストール後は新しいターミナルを開いて再実行すること。

手動でインストールしたい場合は https://www.python.org/ からダウンロード。

### 2. メール認証情報の登録

```
python setup_credentials.py
```

プロンプトに従って:
1. メールプロバイダを選択（iCloud / Gmail / Outlook / Yahoo / カスタム）
2. メールアドレスを入力
3. **アプリパスワード**を入力（通常のログインパスワードではない）

各プロバイダのアプリパスワード発行ページ:
- iCloud: https://appleid.apple.com/
- Gmail: https://myaccount.google.com/apppasswords （2段階認証ON必須）
- Outlook: https://account.microsoft.com/security
- Yahoo!メール: https://account.yahoo.co.jp/

選択結果は `config.json` に書き込まれ、認証情報は Windows Credential Manager に保存される。

### 3. 自動起動の登録

```
install_task.bat を右クリック → 管理者として実行
```

ログオン7秒後にOTP Watcherが起動するタスクスケジューラタスクが登録される。
バッテリー駆動条件も自動で解除される（ノートPCで「キューに挿入済み」になる問題対策）。

### 4. ブラウザ拡張機能のインストール

**Chromium系（Chrome / Brave / Edge / Vivaldi）:**
1. `chrome://extensions/` （Edgeなら `edge://extensions/`）を開く
2. 「デベロッパーモード」をON
3. 「パッケージ化されていない拡張機能を読み込む」→ `extensions/chrome/` フォルダを選択

**Firefox / Floorp:**
- 一時インストール: `about:debugging#/runtime/this-firefox` →「一時的なアドオンを読み込む」→ `extensions/firefox/manifest.json`
- 恒久インストール:
  1. `about:config` → `xpinstall.signatures.required` を `false` に
  2. `extensions/saga-otp-trigger.xpi` をブラウザにドラッグ＆ドロップ

## 動作確認

1. タスクが登録されているか確認:
   ```
   schtasks /query /tn SagaOTP_Watcher /v /fo LIST
   ```
2. プロセスが常駐しているか確認:
   ```
   tasklist | findstr pythonw
   ```
3. HTTPサーバー疎通:
   ```
   curl http://localhost:18924/health
   ```
   → `{"status":"ok"}` が返ればOK
4. SAMLログインを試す → OTPがクリップボードにコピーされ、Toast通知が出る

## トラブルシューティング

### `schtasks /query` の「前回の結果」コード

| コード | 意味 |
|---|---|
| `0` | 成功 |
| `267011` (`0x41303`) | まだ一度も実行されていない（ログオンしていない） |
| `-2147024894` (`0x80070002`) | pythonw.exe が見つからない（WindowsAppsスタブの問題等） |
| `0x41301` | 実行中 |
| `0x41306` | ユーザーが手動で停止した |

### タスクの状態が「キューに挿入済み」

ノートPCのバッテリー駆動条件で起動が保留されている。`install_task.bat` を再実行すれば自動修正される（PowerShell で `DisallowStartIfOnBatteries` / `StopIfGoingOnBatteries` を `false` に）。

GUI で修正する場合: `taskschd.msc` → `SagaOTP_Watcher` → プロパティ → 条件タブ → 電源関連の2項目をOFF。

### 拡張機能からトリガーが届かない

- ログ (`otp_watcher.log`) に「拡張機能からトリガー受信」が出ているか確認
- 出ていない場合:
  - 拡張機能が有効になっているか確認（`chrome://extensions/`）
  - 拡張機能の「サービスワーカー」リンクからDevToolsを開き、エラーを確認
  - `host_permissions` に `http://localhost:18924/*` が含まれているか確認

### `OTPメール待機タイムアウト`

メールが届いていない or プロバイダ設定が違う可能性。`config.json` の `otp_sender_fragment` で送信元を絞っているため、佐大からのメールが指定送信元（デフォルト `LiveCamp`）と一致するか確認。

### ポート 18924 が使えない

別プロセスが占有している。ログに「占有プロセス: ...」が出るので該当プロセスを停止するか、`config.json` の `trigger_port` を別の値に変更（拡張機能側 `manifest.json` の `host_permissions` も合わせて変更要）。

### Toast通知が出ない

- Windows の集中モード（フォーカスアシスト）がONになっていないか確認
- アクションセンターで通知履歴を確認
- 通知音は出るがバナーが出ない場合: 設定 → 通知 → 通知バナーをONに

## アンインストール

1. タスク削除: `uninstall_task.bat` を管理者として実行
2. ブラウザ拡張機能を無効化/削除
3. Windows Credential Manager から `SagaOTP/mail`（および存在すれば `SagaOTP/icloud`）を削除
4. プロセスがまだ動いていれば: `taskkill /im pythonw.exe /f`
5. フォルダごと削除

## 設定リファレンス（config.json）

| キー | デフォルト | 説明 |
|---|---|---|
| `imap_host` | `imap.mail.me.com` | IMAPサーバー（setup_credentials.py で更新される） |
| `imap_port` | `993` | IMAPポート |
| `trigger_port` | `18924` | HTTPトリガー受信ポート（変更時は拡張機能側も修正要） |
| `email_poll_interval` | `1` | メールチェック間隔（秒） |
| `email_poll_timeout` | `1800` | OTPメール待機タイムアウト（秒） |
| `cooldown_seconds` | `60` | 連続トリガー抑止期間（秒） |
| `otp_sender_fragment` | `LiveCamp` | OTPメール送信元の部分一致文字列 |
| `otp_regex` | `ワンタイムパスワード[：:](\d{6,8})` | 本文からOTPを抽出する正規表現 |
| `notification_sound` | `C:\Windows\Media\Alarm06.wav` | （現在は未使用） |

## ライセンス

私的利用向け。

---

# Saga University OTP Watcher

A background tool that automatically copies the one-time password (OTP) — delivered to your registered email during Saga University SAML authentication — to your clipboard.

## How It Works

1. A browser extension detects navigation to `ssoidp.cc.saga-u.ac.jp/idp/profile/SAML2/Redirect`
2. It sends an HTTP POST to a Python script listening on `localhost:18924`
3. The Python script polls your mail server via IMAP and searches for the OTP email
4. The OTP is extracted from the message body, copied to the clipboard, and a Toast notification is shown

Detection latency is **under 0.1 seconds**, thanks to the extension's `webNavigation` API.

## Requirements

- Windows 10 or later (Windows 11 recommended)
- Python 3.9 or later (can be installed automatically via the bundled `python-manager-26.1.msix`)
- Browser: Brave / Chrome / Edge / Firefox / Floorp / etc.
- An IMAP-enabled email account: iCloud / Gmail / Outlook / Yahoo! Mail / any IMAP server

## Setup

### 1. Install Python

If Python is not installed, simply run `install_task.bat` and it will install Python automatically from the bundled `python-manager-26.1.msix`. After installation, open a new terminal and run the batch file again.

To install manually, download from https://www.python.org/.

### 2. Register Your Email Credentials

```
python setup_credentials.py
```

Follow the prompts:
1. Select your mail provider (iCloud / Gmail / Outlook / Yahoo / Custom)
2. Enter your email address
3. Enter your **app password** (not your regular login password)

App password pages by provider:
- iCloud: https://appleid.apple.com/
- Gmail: https://myaccount.google.com/apppasswords (requires 2-Step Verification)
- Outlook: https://account.microsoft.com/security
- Yahoo! Mail: https://account.yahoo.co.jp/

Your selection is written to `config.json` and credentials are stored in Windows Credential Manager.

### 3. Register the Auto-Start Task

```
Right-click install_task.bat → Run as administrator
```

This registers a Task Scheduler task that launches OTP Watcher 7 seconds after logon. Battery-power restrictions are also disabled automatically (prevents the "Queued" status issue on laptops).

### 4. Install the Browser Extension

**Chromium-based browsers (Chrome / Brave / Edge / Vivaldi):**
1. Open `chrome://extensions/` (or `edge://extensions/` for Edge)
2. Enable **Developer mode**
3. Click **Load unpacked** and select the `extensions/chrome/` folder

**Firefox / Floorp:**
- Temporary install: `about:debugging#/runtime/this-firefox` → **Load Temporary Add-on** → select `extensions/firefox/manifest.json`
- Permanent install:
  1. In `about:config`, set `xpinstall.signatures.required` to `false`
  2. Drag and drop `extensions/saga-otp-trigger.xpi` onto the browser window

## Verification

1. Confirm the task is registered:
   ```
   schtasks /query /tn SagaOTP_Watcher /v /fo LIST
   ```
2. Confirm the process is running:
   ```
   tasklist | findstr pythonw
   ```
3. Check the HTTP server:
   ```
   curl http://localhost:18924/health
   ```
   → Should return `{"status":"ok"}`
4. Try a SAML login — the OTP should be copied to your clipboard and a Toast notification should appear.

## Troubleshooting

### Last Run Result codes (`schtasks /query`)

| Code | Meaning |
|---|---|
| `0` | Success |
| `267011` (`0x41303`) | Task has never run (no logon has occurred since registration) |
| `-2147024894` (`0x80070002`) | pythonw.exe not found (WindowsApps stub issue, etc.) |
| `0x41301` | Currently running |
| `0x41306` | Stopped manually by the user |

### Task status shows "Queued"

The task is being held due to the battery power condition. Re-running `install_task.bat` fixes this automatically (it sets `DisallowStartIfOnBatteries` and `StopIfGoingOnBatteries` to `false` via PowerShell).

To fix manually: open `taskschd.msc` → `SagaOTP_Watcher` → Properties → Conditions tab → uncheck both power-related options.

### No trigger received from the extension

- Check `otp_watcher.log` for the message "拡張機能からトリガー受信" (trigger received from extension)
- If it is not appearing:
  - Confirm the extension is enabled (`chrome://extensions/`)
  - Open DevTools from the extension's service worker link and check for errors
  - Confirm `host_permissions` includes `http://localhost:18924/*`

### OTP email wait timeout

The email may not have arrived, or the provider settings may be incorrect. The search is filtered by `otp_sender_fragment` in `config.json` (default: `LiveCamp`). Verify that OTP emails from the university actually match this sender fragment.

### Port 18924 unavailable

Another process is occupying the port. The log will show something like "占有プロセス: ..." (owning process: ...) to help identify it. Either stop that process or change `trigger_port` in `config.json` — and update `host_permissions` in the extension's `manifest.json` to match.

### Toast notifications not appearing

- Check that Focus Assist (Do Not Disturb) is not enabled in Windows
- Check the Action Center notification history
- If sound plays but no banner appears: Settings → Notifications → enable notification banners

## Uninstall

1. Remove the task: run `uninstall_task.bat` as administrator
2. Disable or remove the browser extension
3. Delete `SagaOTP/mail` (and `SagaOTP/icloud` if present) from Windows Credential Manager
4. If the process is still running: `taskkill /im pythonw.exe /f`
5. Delete the project folder

## Configuration Reference (`config.json`)

| Key | Default | Description |
|---|---|---|
| `imap_host` | `imap.mail.me.com` | IMAP server hostname (updated by `setup_credentials.py`) |
| `imap_port` | `993` | IMAP port |
| `trigger_port` | `18924` | HTTP trigger listener port (changing this requires updating the extension too) |
| `email_poll_interval` | `1` | Email polling interval in seconds |
| `email_poll_timeout` | `1800` | Timeout for waiting for the OTP email, in seconds |
| `cooldown_seconds` | `60` | Suppression window after an OTP is copied, in seconds |
| `otp_sender_fragment` | `LiveCamp` | Partial match string for the OTP email sender |
| `otp_regex` | `ワンタイムパスワード[：:](\d{6,8})` | Regex to extract the OTP from the message body |
| `notification_sound` | `C:\Windows\Media\Alarm06.wav` | (currently unused) |

## License

For personal use.
