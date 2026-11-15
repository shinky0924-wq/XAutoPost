#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chrome カスタムプロファイル起動ツール (chrome_launcher.py / chrome_launcher.exe)

- 指定したプロファイル名で browser_sessions/ 配下に独立したChromeプロファイルを作成・起動します
- サイト（X等）は勝手に開きません（完全手動アクセス用）
- --no-first-run を付けないため、初回作成時は「Chrome にログイン」画面が表示されます
- ログイン後のCookie・セッションはそのまま保持され、2回目以降は X_AutoTweetBot が自動で引き継ぎます
"""

import os
import re
import sys
import shutil
import subprocess

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def get_app_dir() -> str:
    """EXE実行時・スクリプト実行時ともに実行ファイルと同じフォルダを基準にする"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def find_chrome_executable() -> str | None:
    """システムにインストールされている Google Chrome の実行ファイルパスを検出"""
    candidates = []

    if sys.platform == "win32":
        for env_key in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env_key)
            if base:
                candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))

        try:
            import winreg
            reg_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
            ]
            for root, subkey in reg_paths:
                try:
                    with winreg.OpenKey(root, subkey) as key:
                        val, _ = winreg.QueryValueEx(key, "")
                        if val:
                            candidates.append(val)
                except OSError:
                    pass
        except ImportError:
            pass

        for cmd_name in ("chrome.exe", "chrome"):
            found = shutil.which(cmd_name)
            if found:
                candidates.append(found)
    elif sys.platform == "darwin":
        candidates.append("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
        found = shutil.which("google-chrome")
        if found:
            candidates.append(found)
    else:
        for cmd_name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            found = shutil.which(cmd_name)
            if found:
                candidates.append(found)

    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def sanitize_profile_name(raw_name: str) -> str:
    """プロファイル名から先頭の@やWindowsで禁止されている文字を除去"""
    cleaned = raw_name.strip().lstrip("@").strip()
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", cleaned)
    return cleaned


def resolve_profile_dir(sessions_root: str, clean_name: str) -> tuple[str, bool]:
    """
    プロファイル保存先パスを解決する（auto_tweet_playwright.py と共通ルール）
    Returns: (profile_dir_path, already_exists)
    """
    exact_dir = os.path.join(sessions_root, clean_name)
    underscore_dir = os.path.join(sessions_root, f"_{clean_name}")

    if os.path.exists(exact_dir):
        return os.path.abspath(exact_dir), True
    if os.path.exists(underscore_dir):
        return os.path.abspath(underscore_dir), True

    os.makedirs(exact_dir, exist_ok=True)
    return os.path.abspath(exact_dir), False


def list_existing_profiles(sessions_root: str) -> list[str]:
    """既存のプロファイル一覧を取得"""
    if not os.path.isdir(sessions_root):
        return []
    profiles = []
    for entry in sorted(os.listdir(sessions_root)):
        full = os.path.join(sessions_root, entry)
        if os.path.isdir(full):
            profiles.append(entry)
    return profiles


def launch_profile():
    app_dir = get_app_dir()
    sessions_root = os.path.join(app_dir, "browser_sessions")
    os.makedirs(sessions_root, exist_ok=True)

    chrome_exe = find_chrome_executable()
    if not chrome_exe:
        print("\n❌ Google Chrome が見つかりませんでした。")
        print("Google Chrome をインストールしてから再度実行してください。")
        input("\nEnterキーを押すと終了します...")
        return

    while True:
        print("\n====================================================")
        print("Chrome カスタムプロファイル起動ツール")
        print("====================================================")
        print(f"保存先フォルダ: {sessions_root}")
        print("※ サイトは自動で開きません。手動で操作・ログインしてください。")

        existing = list_existing_profiles(sessions_root)
        if existing:
            print(f"既存プロファイル ({len(existing)}件): {', '.join(existing)}")

        print()
        raw_input_name = input("作成・使用するプロファイル名を入力してください: ").strip()
        if not raw_input_name:
            print("⚠️ プロファイル名が空です。終了します。")
            break

        clean_name = sanitize_profile_name(raw_input_name)
        if not clean_name:
            print("⚠️ 有効なプロファイル名を入力してください。")
            continue

        profile_dir, already_exists = resolve_profile_dir(sessions_root, clean_name)

        if already_exists:
            print(f"\n📂 既存プロファイルを読み込みます: {profile_dir}")
        else:
            print(f"\n✨ 新規プロファイルを作成します: {profile_dir}")

        print("🚀 Google Chrome を起動しています...\n")
        print("【手動登録の手順】")
        print("  1. 「Chrome にログイン」画面が表示されたら操作してください（不要なら「ログアウト状態を保持」）")
        print("  2. ご自身で X などのサイトへアクセスし、手動でログインしてください")
        print("  3. ログインが完了したら、開いた Chrome を「×」で閉じてください")
        print("  （※ 次回からはログイン情報が保持され、自動投稿ツールがそのまま利用できます）\n")

        # URL指定なし・--no-first-run なしで起動することで、
        # 初回は通常の「Chrome にログイン」画面が開き、勝手にサイトを開きません
        cmd = [
            chrome_exe,
            f"--user-data-dir={profile_dir}",
        ]

        try:
            proc = subprocess.Popen(cmd)
            print("⏳ Chrome ウィンドウが開いています（Chrome を閉じると次に進みます）...")
            proc.wait()
            print(f"\n✅ プロファイル [{os.path.basename(profile_dir)}] のセッション情報を保存しました。")
        except Exception as e:
            print(f"\n❌ Chrome の起動中にエラーが発生しました: {e}")

        again = input("\n別のプロファイルを続けて起動しますか？ (y / Enterで終了): ").strip().lower()
        if again not in ("y", "yes"):
            break


def main():
    try:
        launch_profile()
    except KeyboardInterrupt:
        print("\n中断しました。")
    except Exception as e:
        print(f"\n予期せぬエラーが発生しました: {e}")
        input("Enterキーを押すと終了します...")


if __name__ == "__main__":
    main()
