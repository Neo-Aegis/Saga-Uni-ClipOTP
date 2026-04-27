"""
メール認証情報をWindows Credential Managerに保存し、
選択したIMAPプロバイダ設定を config.json に書き込むセットアップスクリプト。
初回および認証情報変更時に実行する。
"""

import ctypes
import ctypes.wintypes as wintypes
import getpass
import imaplib
import json
import os
import sys

CREDENTIAL_TARGET = "SagaOTP/mail"
LEGACY_CREDENTIAL_TARGET = "SagaOTP/icloud"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")
IMAP_TIMEOUT_SECONDS = 30

# (表示名, IMAPホスト, ポート, アプリパスワード発行URL)
PROVIDERS = [
    ("iCloud",          "imap.mail.me.com",         993, "https://appleid.apple.com/"),
    ("Gmail",           "imap.gmail.com",           993, "https://myaccount.google.com/apppasswords"),
    ("Outlook/Hotmail", "outlook.office365.com",    993, "https://account.microsoft.com/security"),
    ("Yahoo!メール",    "imap.mail.yahoo.co.jp",    993, "https://account.yahoo.co.jp/"),
]

# --- Windows Credential Manager API ---

CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2

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

def write_credential(target: str, username: str, password: str) -> None:
    password_bytes = password.encode("utf-16-le")
    blob = (ctypes.c_ubyte * len(password_bytes))(*password_bytes)

    cred = CREDENTIAL()
    cred.Flags = 0
    cred.Type = CRED_TYPE_GENERIC
    cred.TargetName = target
    cred.Comment = "SagaOTP mail credentials"
    cred.CredentialBlobSize = len(password_bytes)
    cred.CredentialBlob = blob
    cred.Persist = CRED_PERSIST_LOCAL_MACHINE
    cred.AttributeCount = 0
    cred.Attributes = None
    cred.TargetAlias = None
    cred.UserName = username

    if not advapi32.CredWriteW(ctypes.byref(cred), 0):
        raise ctypes.WinError(ctypes.get_last_error())


def read_credential(target: str) -> tuple[str, str] | None:
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


def delete_credential(target: str) -> bool:
    return bool(advapi32.CredDeleteW(target, CRED_TYPE_GENERIC, 0))


# --- プロバイダ選択 ---

def choose_provider() -> tuple[str, int, str]:
    """プロバイダ選択メニューを表示し (host, port, app_password_url) を返す。"""
    print()
    print("メールプロバイダを選択してください:")
    for i, (name, host, port, _) in enumerate(PROVIDERS, start=1):
        print(f"  {i}) {name}  ({host}:{port})")
    print(f"  {len(PROVIDERS) + 1}) カスタム (ホスト/ポートを手入力)")
    print()

    while True:
        ans = input(f"番号を選択 [1-{len(PROVIDERS) + 1}]: ").strip()
        if not ans.isdigit():
            print("  数字を入力してください。")
            continue
        idx = int(ans)
        if 1 <= idx <= len(PROVIDERS):
            _, host, port, url = PROVIDERS[idx - 1]
            return host, port, url
        if idx == len(PROVIDERS) + 1:
            host = input("  IMAPホスト名: ").strip()
            if not host:
                print("  ホスト名が空です。")
                continue
            port_str = input("  ポート番号 [993]: ").strip() or "993"
            if not port_str.isdigit():
                print("  ポートは数字で入力してください。")
                continue
            return host, int(port_str), ""
        print(f"  1〜{len(PROVIDERS) + 1}の範囲で入力してください。")


# --- IMAP接続テスト ---

def test_imap_connection(host: str, port: int, email_addr: str, app_password: str) -> bool:
    print(f"\n{host}:{port} に接続テスト中...")
    try:
        imap = imaplib.IMAP4_SSL(host, port, timeout=IMAP_TIMEOUT_SECONDS)
        imap.login(email_addr, app_password)
        status, _ = imap.select("INBOX", readonly=True)
        if status == "OK":
            print("  接続成功! INBOXを正常に開けました。")
            try:
                imap.logout()
            except Exception:
                pass
            return True
        else:
            print(f"  INBOXの選択に失敗: {status}")
            try:
                imap.logout()
            except Exception:
                pass
            return False
    except imaplib.IMAP4.error as e:
        print(f"  IMAP認証エラー: {e}")
        print("  メールアドレスまたはアプリパスワードを確認してください。")
        return False
    except Exception as e:
        print(f"  接続エラー: {e}")
        return False


# --- config.json更新 ---

def update_config_imap(host: str, port: int) -> None:
    """config.json の imap_host / imap_port を更新する。他キーは保持。"""
    config = {}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            print(f"  既存config.json読み込み失敗 ({e})。新規作成します。")
            config = {}
    config["imap_host"] = host
    config["imap_port"] = port
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"  config.json を更新しました: imap_host={host}, imap_port={port}")


# --- メイン処理 ---

def main():
    print("=" * 50)
    print("佐大OTP - メール認証情報セットアップ")
    print("=" * 50)

    # 既存認証情報の確認（新キー優先、旧キーもチェック）
    existing = read_credential(CREDENTIAL_TARGET)
    legacy = read_credential(LEGACY_CREDENTIAL_TARGET)
    if existing:
        print(f"\n既存の認証情報が見つかりました: {existing[0]}")
        ans = input("上書きしますか？ (y/N): ").strip().lower()
        if ans != "y":
            print("キャンセルしました。")
            return
    elif legacy:
        print(f"\n旧バージョンの認証情報が見つかりました: {legacy[0]}")
        print("新しい形式に移行します。")

    # プロバイダ選択
    host, port, app_password_url = choose_provider()

    # メールアドレス入力
    print()
    email_addr = input("メールアドレス: ").strip()
    if not email_addr:
        print("メールアドレスが空です。終了します。")
        sys.exit(1)

    # アプリパスワード入力
    print()
    print("アプリパスワード（プロバイダで発行した16〜20文字程度のパスワード）を入力してください。")
    if app_password_url:
        print(f"発行ページ: {app_password_url}")
    print("通常のログインパスワードでは認証できないことが多いので注意。")
    app_password = getpass.getpass("アプリパスワード: ").strip()
    if not app_password:
        print("パスワードが空です。終了します。")
        sys.exit(1)

    # 接続テスト
    if not test_imap_connection(host, port, email_addr, app_password):
        print()
        ans = input("接続テストに失敗しましたが、保存しますか？ (y/N): ").strip().lower()
        if ans != "y":
            print("キャンセルしました。")
            sys.exit(1)

    # 保存
    try:
        write_credential(CREDENTIAL_TARGET, email_addr, app_password)
        print(f"\n認証情報を保存しました（ターゲット: {CREDENTIAL_TARGET}）")
    except Exception as e:
        print(f"認証情報の保存エラー: {e}")
        sys.exit(1)

    update_config_imap(host, port)

    # 旧キーが残っていれば削除を提案
    if legacy and not existing:
        ans = input(f"\n旧キー ({LEGACY_CREDENTIAL_TARGET}) を削除しますか？ (Y/n): ").strip().lower()
        if ans != "n":
            if delete_credential(LEGACY_CREDENTIAL_TARGET):
                print("旧キーを削除しました。")
            else:
                print("旧キーの削除に失敗しました（手動で削除してください）。")

    print("\nセットアップ完了!")


if __name__ == "__main__":
    main()
