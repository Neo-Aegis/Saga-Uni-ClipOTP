"""
佐大OTP Watcher v2 - ブラウザ拡張機能からのHTTPトリガーでSAML URL検知し、
iCloudからOTPを取得してクリップボードにコピーする。
バックグラウンドで常駐動作する (.pyw)。
"""

import base64
import ctypes
import ctypes.wintypes as wintypes
import email as email_mod
import email.utils
import imaplib
import json
import logging
import logging.handlers
import os
import re
import signal
import subprocess
import sys
import threading
import time
import winsound
from datetime import datetime, timezone, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler

# --- 定数 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")
LOG_PATH = os.path.join(SCRIPT_DIR, "otp_watcher.log")
MUTEX_NAME = "SagaOTP_Watcher_SingleInstance_Mutex"
IMAP_TIMEOUT_SECONDS = 30
IMAP_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# デフォルト設定
DEFAULTS = {
    "imap_host": "imap.mail.me.com",
    "imap_port": 993,
    "trigger_port": 18924,
    "email_poll_interval": 1,
    "email_poll_timeout": 1800,
    "cooldown_seconds": 60,
    "otp_sender_fragment": "LiveCamp",
    "otp_regex": r"ワンタイムパスワード[：:](\d{6,8})",
    "notification_sound": r"C:\Windows\Media\Alarm06.wav",
}

# --- ロガー ---
logger = logging.getLogger("SagaOTP")
logger.setLevel(logging.INFO)
handler = logging.handlers.RotatingFileHandler(
    LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
)
handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(handler)


# --- 設定の読み込み ---
def load_config() -> dict:
    config = dict(DEFAULTS)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                user_config = json.load(f)
            config.update(user_config)
        except Exception as e:
            logger.warning("config.json読み込みエラー: %s (デフォルト値を使用)", e)
    return config


# --- Windows Credential Manager ---
CRED_TYPE_GENERIC = 1


class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", wintypes.FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.c_void_p),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


advapi32 = ctypes.windll.advapi32


def read_credential(target: str) -> tuple[str, str] | None:
    """Windows Credential Managerから認証情報を読み取る。"""
    pcred = ctypes.POINTER(CREDENTIAL)()
    if not advapi32.CredReadW(target, CRED_TYPE_GENERIC, 0, ctypes.byref(pcred)):
        return None
    try:
        cred = pcred.contents
        username = cred.UserName or ""
        blob_size = cred.CredentialBlobSize
        password = ctypes.string_at(cred.CredentialBlob, blob_size).decode("utf-16-le")
        return username, password
    finally:
        advapi32.CredFree(pcred)


# --- HTTPトリガーサーバー ---

trigger_event = threading.Event()


class TriggerHandler(BaseHTTPRequestHandler):
    """ブラウザ拡張機能からのトリガーを受け取るHTTPハンドラー。"""

    def do_POST(self):
        if self.path == "/trigger":
            trigger_event.set()
            logger.info("拡張機能からトリガー受信。")
            self._respond(200, {"status": "triggered"})
        else:
            self._respond(404, {"error": "not found"})

    def do_GET(self):
        if self.path == "/health":
            self._respond(200, {"status": "ok"})
        else:
            self._respond(404, {"error": "not found"})

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.end_headers()

    def _respond(self, code: int, body: dict):
        self.send_response(code)
        self._cors_headers()
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(body).encode())

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def log_message(self, format, *args):
        # デフォルトのstderrログを抑制
        pass


def start_trigger_server(port: int):
    """HTTPトリガーサーバーをデーモンスレッドで起動する。"""
    server = HTTPServer(("127.0.0.1", port), TriggerHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server


def diagnose_port_owner(port: int) -> str:
    """指定ポートを占有しているプロセス名/PIDを返す（特定不能なら空文字）。"""
    ps_script = (
        f"$c = Get-NetTCPConnection -LocalPort {port} -ErrorAction SilentlyContinue | "
        "Select-Object -First 1; "
        "if ($c) { $p = Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue; "
        "if ($p) { Write-Output (\"$($p.ProcessName) (PID $($p.Id))\") } }"
    )
    try:
        encoded = base64.b64encode(ps_script.encode("utf-16-le")).decode("ascii")
        result = subprocess.run(
            ["powershell", "-NoProfile", "-EncodedCommand", encoded],
            capture_output=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return result.stdout.decode("utf-8", errors="replace").strip()
    except Exception:
        return ""


# --- IMAPメール取得 ---

def get_email_body(msg) -> str:
    """メールメッセージから本文テキストを抽出する。"""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                charset = part.get_content_charset() or "utf-8"
                try:
                    return part.get_payload(decode=True).decode(charset, errors="replace")
                except Exception:
                    return part.get_payload(decode=True).decode("utf-8", errors="replace")
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            return msg.get_payload(decode=True).decode(charset, errors="replace")
        except Exception:
            return msg.get_payload(decode=True).decode("utf-8", errors="replace")
    return ""


def connect_imap(config: dict, email_addr: str, app_password: str):
    """IMAPサーバに接続してINBOXを開く。接続済みのimapオブジェクトを返す。"""
    imap = imaplib.IMAP4_SSL(
        config["imap_host"], config["imap_port"], timeout=IMAP_TIMEOUT_SECONDS
    )
    imap.login(email_addr, app_password)
    imap.select("INBOX", readonly=True)
    return imap


def search_otp_on_connection(imap, config: dict, since_time: datetime) -> str | None:
    """既存のIMAP接続上でOTPメールを検索する。見つからなければNone。"""
    # NOOPでサーバー側の状態を更新（新着メール反映）
    imap.noop()

    # サーバーサイドでFROMとSINCEを絞り込み（全メール走査を回避）
    # IMAP仕様の英語月名を固定で組み立て（ロケール非依存）
    date_str = f"{since_time.day:02d}-{IMAP_MONTHS[since_time.month - 1]}-{since_time.year:04d}"
    sender_fragment = config["otp_sender_fragment"]
    status, data = imap.search(
        None, f'(SINCE {date_str} FROM "{sender_fragment}")'
    )

    if status != "OK" or not data[0]:
        return None

    msg_ids = data[0].split()
    otp_regex = re.compile(config.get("otp_regex", DEFAULTS["otp_regex"]))

    # 最新のメールから順にチェック（最新1-2件で十分）
    for msg_id in reversed(msg_ids[-3:]):
        status, full_data = imap.fetch(msg_id, "(BODY[])")
        if status != "OK" or not full_data or not isinstance(full_data[0], tuple):
            continue

        msg = email_mod.message_from_bytes(full_data[0][1])

        # メール日時チェック
        email_date_str = msg.get("Date", "")
        try:
            email_date = email_mod.utils.parsedate_to_datetime(email_date_str)
            if email_date.tzinfo is None:
                email_date = email_date.replace(tzinfo=timezone.utc)
            if email_date < since_time:
                continue
        except Exception:
            pass

        # 本文からOTP抽出
        body = get_email_body(msg)
        match = otp_regex.search(body)
        if match:
            return match.group(1)

    return None


# --- クリップボード ---

def copy_to_clipboard(text: str) -> None:
    """テキストをWindowsクリップボードにコピーする。"""
    process = subprocess.Popen(
        ["clip.exe"],
        stdin=subprocess.PIPE,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    process.communicate(input=text.encode("utf-8"))


# --- 通知 ---

def show_toast(title: str, message: str) -> None:
    """Windows Toast通知を表示する（PowerShell EncodedCommand経由でロケール非依存）。"""
    title_safe = title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    msg_safe = message.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    ps_script = (
        "[Windows.UI.Notifications.ToastNotificationManager, "
        "Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null; "
        "[Windows.Data.Xml.Dom.XmlDocument, "
        "Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null; "
        f"$t = '<toast><visual><binding template=\"ToastText02\">"
        f"<text id=\"1\">{title_safe}</text>"
        f"<text id=\"2\">{msg_safe}</text>"
        f"</binding></visual></toast>'; "
        "$x = New-Object Windows.Data.Xml.Dom.XmlDocument; "
        "$x.LoadXml($t); "
        "$n = [Windows.UI.Notifications.ToastNotification]::new($x); "
        "[Windows.UI.Notifications.ToastNotificationManager]"
        "::CreateToastNotifier('SagaOTP').Show($n)"
    )
    try:
        # PowerShellの-EncodedCommandはUTF-16-LE+Base64を要求する。
        # これによりargs変換時のCP932依存を排除し、非日本語Windowsでも文字化けしない。
        encoded = base64.b64encode(ps_script.encode("utf-16-le")).decode("ascii")
        subprocess.Popen(
            ["powershell", "-NoProfile", "-EncodedCommand", encoded],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        logger.debug("Toast通知エラー: %s", e)


def notify_user(otp: str, config: dict) -> None:
    """OTP取得完了を通知する（Toastのみ）。"""
    show_toast("佐大OTP", f"ワンタイムパスワードをコピーしました: {otp}")


# --- 名前付きMutexによる多重起動防止 ---

ERROR_ALREADY_EXISTS = 183
_mutex_handle = None


def acquire_single_instance_lock() -> bool:
    """名前付きMutexを取得する。既に他インスタンスが保持していればFalse。"""
    global _mutex_handle
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        return False
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(handle)
        return False
    _mutex_handle = handle
    return True


def release_single_instance_lock():
    """Mutexハンドルを解放する（プロセス終了時にOSも自動解放するが明示的に閉じる）。"""
    global _mutex_handle
    if _mutex_handle:
        ctypes.windll.kernel32.CloseHandle(_mutex_handle)
        _mutex_handle = None


# --- メインループ ---

def main():
    if not acquire_single_instance_lock():
        logger.info("既に別のインスタンスが実行中です。終了します。")
        sys.exit(0)

    def shutdown_handler(signum, frame):
        logger.info("シャットダウンシグナル受信 (%s)。終了します。", signum)
        release_single_instance_lock()
        sys.exit(0)

    signal.signal(signal.SIGTERM, shutdown_handler)
    signal.signal(signal.SIGINT, shutdown_handler)

    config = load_config()

    # 新キー (SagaOTP/mail) を優先、旧キー (SagaOTP/icloud) を後方互換でフォールバック
    cred = read_credential("SagaOTP/mail") or read_credential("SagaOTP/icloud")
    if not cred:
        logger.error("認証情報が見つかりません。setup_credentials.py を実行してください。")
        show_toast("佐大OTP エラー", "認証情報が見つかりません。setup_credentials.py を実行してください。")
        release_single_instance_lock()
        sys.exit(1)

    email_addr, app_password = cred

    # HTTPトリガーサーバー起動
    port = config.get("trigger_port", DEFAULTS["trigger_port"])
    try:
        server = start_trigger_server(port)
        logger.info("OTP Watcher v2 起動 (メール: %s, ポート: %d)", email_addr, port)
    except OSError as e:
        owner = diagnose_port_owner(port)
        if owner:
            logger.error("HTTPサーバー起動失敗 (ポート %d): %s | 占有プロセス: %s", port, e, owner)
            show_toast("佐大OTP エラー", f"ポート{port}を別プロセスが占有: {owner}")
        else:
            logger.error("HTTPサーバー起動失敗 (ポート %d): %s", port, e)
            show_toast("佐大OTP エラー", f"ポート{port}が利用できません: {e}")
        release_single_instance_lock()
        sys.exit(1)

    state = "IDLE"
    last_otp_time = 0.0
    saml_trigger_time = None
    poll_start = 0.0
    imap_conn = None
    last_imap_keepalive = 0.0
    IMAP_KEEPALIVE_INTERVAL = 270
    # IMAP接続失敗の連続回数とtoast発火済みフラグ（デバウンス用）
    imap_fail_count = 0
    imap_fail_toast_shown = False
    IMAP_FAIL_TOAST_THRESHOLD = 3

    def close_imap():
        nonlocal imap_conn
        if imap_conn:
            try:
                imap_conn.logout()
            except Exception:
                pass
            imap_conn = None

    def ensure_imap():
        nonlocal imap_conn, last_imap_keepalive, imap_fail_count, imap_fail_toast_shown
        now = time.time()

        if imap_conn:
            if now - last_imap_keepalive > IMAP_KEEPALIVE_INTERVAL:
                try:
                    imap_conn.noop()
                    last_imap_keepalive = now
                except Exception:
                    logger.info("IMAP接続切断を検知。再接続します。")
                    imap_conn = None

        if not imap_conn:
            try:
                imap_conn = connect_imap(config, email_addr, app_password)
                last_imap_keepalive = now
                logger.info("IMAP接続確立。")
                # 復旧したのでデバウンスフラグをリセット
                imap_fail_count = 0
                imap_fail_toast_shown = False
            except Exception as e:
                logger.error("IMAP接続失敗: %s", e)
                imap_conn = None
                imap_fail_count += 1
                if imap_fail_count >= IMAP_FAIL_TOAST_THRESHOLD and not imap_fail_toast_shown:
                    show_toast(
                        "佐大OTP エラー",
                        f"メールサーバーへの接続に{imap_fail_count}回連続で失敗しています。",
                    )
                    imap_fail_toast_shown = True

    ensure_imap()

    try:
        while True:
            if state == "IDLE":
                # イベント待ち（タイムアウト1秒でキープアライブも実行）
                triggered = trigger_event.wait(timeout=1)

                if triggered:
                    trigger_event.clear()

                    # クールダウン中はスキップ
                    if time.time() - last_otp_time < config["cooldown_seconds"]:
                        logger.info("クールダウン中のためトリガーをスキップ。")
                        continue

                    logger.info("SAML URLへのアクセスを検知。メールポーリングを開始します。")
                    saml_trigger_time = datetime.now(timezone.utc) - timedelta(seconds=30)
                    state = "POLLING_EMAIL"
                    poll_start = time.time()

                    if not imap_conn:
                        ensure_imap()
                    if not imap_conn:
                        # ensure_imap内でデバウンス済みtoastが出るのでここでは追加発火しない
                        state = "IDLE"
                else:
                    # タイムアウト: キープアライブ
                    ensure_imap()

            elif state == "POLLING_EMAIL":
                try:
                    otp = search_otp_on_connection(imap_conn, config, saml_trigger_time)

                    if otp:
                        copy_to_clipboard(otp)
                        notify_user(otp, config)
                        logger.info("OTPをクリップボードにコピーしました: %s****", otp[:2])
                        last_otp_time = time.time()
                        state = "IDLE"
                    elif time.time() - poll_start > config["email_poll_timeout"]:
                        logger.warning("OTPメール待機タイムアウト (%d秒)", config["email_poll_timeout"])
                        state = "IDLE"
                    else:
                        time.sleep(config["email_poll_interval"])

                except Exception as e:
                    logger.warning("IMAPポーリングエラー: %s。再接続します。", e)
                    close_imap()
                    ensure_imap()
                    if not imap_conn:
                        logger.error("IMAP再接続失敗。IDLEに戻ります。")
                        # ensure_imap内でデバウンス済みtoastが出るのでここでは追加発火しない
                        state = "IDLE"

    except KeyboardInterrupt:
        logger.info("KeyboardInterruptにより終了。")
    finally:
        server.shutdown()
        close_imap()
        release_single_instance_lock()
        logger.info("OTP Watcher 終了。")


if __name__ == "__main__":
    main()
