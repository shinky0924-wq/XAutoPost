#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
X Auto Tweet Bot コントロールパネル (完全統合 GUI)
- 稼働状況ダッシュボード（アカウント×テーマ進捗・次回予定カウントダウン・1件テスト・時刻シミュレーション・休止/除外管理）
- テーマ・ポスト直編集（Excelを開かずにアカウント・テーマ別の設定変更 / ポスト編集 / 一括リセット / 新規作成）
- Chromeプロファイル管理（作成 / 一覧 / リネーム / 削除 / 手動Chrome完全分離起動 / ログイン状態一括チェック）
- アカウント情報.xlsx 管理
- 画像＆クールダウン管理（image_history.json 可視化・リセット）
- アカウント別ログ（log/<アカウント>.log）＆スクリーンショット（screenshots/<アカウント>/）ビューア
"""

import os
import re
import sys
import glob
import json
import time
import queue
import shutil
import threading
import subprocess
import unicodedata
from datetime import datetime, time as dt_time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog, filedialog

try:
    import openpyxl
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

try:
    import update_themes_excel as theme_helper
    HAS_THEME_HELPER = True
except Exception:
    theme_helper = None
    HAS_THEME_HELPER = False


LEGACY_IGNORED_SHEETS = {"設定", "日常ツイート"}
IMAGE_COOLDOWN_SAME_DAYS = 30
IMAGE_COOLDOWN_CROSS_DAYS = 7


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def normalize_text(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    if isinstance(val, dt_time):
        return val.strftime("%H:%M")
    return unicodedata.normalize("NFKC", str(val)).strip()


def sanitize_profile_name(raw_name: str) -> str:
    cleaned = str(raw_name or "").strip().lstrip("@").strip()
    return re.sub(r'[<>:"/\\|?*]', "_", cleaned)


def parse_single_time_to_minutes(token: str) -> int | None:
    s = normalize_text(token)
    if not s:
        return None
    s = re.sub(r'\s+', '', s)
    is_pm = ("午後" in s or "pm" in s.lower())
    s_clean = re.sub(r'(翌|午前|午後|am|pm)', '', s, flags=re.IGNORECASE)
    m = re.match(r'^(\d{1,2})\s*(?::(\d{1,2})|時(?:(半)|(\d{1,2})分?)?)?$', s_clean)
    if not m:
        return None
    hour = int(m.group(1))
    if m.group(2) is not None:
        minute = int(m.group(2))
    elif m.group(3) == "半":
        minute = 30
    elif m.group(4) is not None:
        minute = int(m.group(4))
    else:
        minute = 0
    if is_pm and 1 <= hour <= 11:
        hour += 12
    if hour == 24 and minute == 0:
        return 24 * 60
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour * 60 + minute


def is_time_in_ranges(time_spec, now_dt: datetime | None = None) -> bool:
    raw = normalize_text(time_spec)
    if not raw or raw.lower() in ("none", "nan", "終日", "24時間", "いつでも", "常時", "all"):
        return True
    if now_dt is None:
        now_dt = datetime.now()
    cur_min = now_dt.hour * 60 + now_dt.minute
    segments = [seg.strip() for seg in re.split(r'[,、;\n/]+', raw) if seg.strip()]
    if not segments:
        return True
    valid_range_found = False
    for seg in segments:
        unified = re.sub(r'(?:から|〜|～|~|−|–|—|ー|-|\bto\b)', '~', seg, flags=re.IGNORECASE)
        parts = [p.strip() for p in unified.split('~') if p.strip()]
        if len(parts) < 2:
            single_m = parse_single_time_to_minutes(parts[0]) if parts else None
            if single_m is not None:
                valid_range_found = True
                if cur_min == single_m:
                    return True
            continue
        s_min = parse_single_time_to_minutes(parts[0])
        e_min = parse_single_time_to_minutes(parts[1])
        if s_min is None or e_min is None:
            continue
        valid_range_found = True
        if s_min == 0 and e_min == 0:
            return True
        if s_min <= e_min:
            if s_min <= cur_min <= e_min:
                return True
        else:
            if cur_min >= s_min or cur_min <= e_min:
                return True
    return False if valid_range_found else True


def is_today_allowed(weekday_spec, now_dt: datetime | None = None) -> bool:
    raw = normalize_text(weekday_spec)
    if not raw or raw.lower() in ("毎日", "全日", "すべて", "all", "*", "none", "nan"):
        return True
    if now_dt is None:
        now_dt = datetime.now()
    wd = now_dt.weekday()
    if "平日" in raw or "月~金" in raw or "月-金" in raw or "月〜金" in raw:
        return wd < 5
    if "土日" in raw or "週末" in raw:
        return wd >= 5
    jp_days = ["月", "火", "水", "木", "金", "土", "日"]
    en_days = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
    low = raw.lower()
    matched = any(ch in raw for ch in jp_days) or any(ed in low for ed in en_days)
    if not matched:
        return True
    return (jp_days[wd] in raw) or (en_days[wd] in low)


def find_chrome_executable() -> str | None:
    candidates = []
    if sys.platform == "win32":
        for env_key in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env_key)
            if base:
                candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
        try:
            import winreg
            for root, subkey in (
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
            ):
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
    else:
        for cmd_name in ("google-chrome", "chromium", "chromium-browser"):
            found = shutil.which(cmd_name)
            if found:
                candidates.append(found)

    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None


def find_python_cli() -> str:
    exe = sys.executable
    if exe and os.path.basename(exe).lower() == "pythonw.exe":
        py_cli = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.isfile(py_cli):
            return py_cli
    return exe or "python"


def kill_external_bots_on_windows() -> int:
    if sys.platform != "win32":
        return 0
    ps_script = (
        "$cnt = 0; "
        "$procs = Get-CimInstance Win32_Process | Where-Object { "
        "  $_.Name -eq 'X_AutoTweetBot.exe' -or "
        "  (($_.Name -eq 'python.exe' -or $_.Name -eq 'pythonw.exe') -and $_.CommandLine -like '*auto_tweet_playwright.py*') "
        "}; "
        "foreach ($p in $procs) { "
        "  try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue; $cnt++ } catch {} "
        "}; "
        "Write-Output $cnt"
    )
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        out = (res.stdout or "").strip().splitlines()
        return int(out[-1]) if out and out[-1].isdigit() else 0
    except Exception:
        return 0


def is_pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        try:
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
            if not handle:
                return False
            try:
                exit_code = ctypes.c_ulong()
                if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                    return exit_code.value == STILL_ACTIVE
                return True
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return False
    try:
        os.kill(pid, 0)
        return True
    except Exception:
        return False


def check_any_bot_running_fast(app_dir: str) -> bool:
    pid_file = os.path.join(app_dir, "bot_running.pid")
    if not os.path.exists(pid_file):
        return False
    try:
        with open(pid_file, "r", encoding="utf-8") as f:
            pid = int(f.read().strip())
        if is_pid_alive(pid):
            return True
    except Exception:
        pass
    try:
        os.remove(pid_file)
    except Exception:
        pass
    return False


def request_graceful_stop(app_dir: str) -> None:
    flag_path = os.path.join(app_dir, "stop_bot.flag")
    try:
        with open(flag_path, "w", encoding="utf-8") as f:
            f.write("stop")
    except Exception:
        pass


def get_active_bot_profile(app_dir: str) -> str:
    p = os.path.join(app_dir, "bot_active_profile.txt")
    if not os.path.exists(p):
        return ""
    try:
        with open(p, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def load_paused_set(app_dir: str) -> set[str]:
    p = os.path.join(app_dir, "paused_accounts.json")
    if not os.path.exists(p):
        return set()
    try:
        with open(p, "r", encoding="utf-8") as f:
            arr = json.load(f)
        if isinstance(arr, list):
            return {sanitize_profile_name(x).lower() for x in arr if str(x).strip()}
    except Exception:
        pass
    return set()


def save_paused_set(app_dir: str, paused: set[str]) -> None:
    p = os.path.join(app_dir, "paused_accounts.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(sorted(paused), f, ensure_ascii=False, indent=2)


def load_suspended_entries(app_dir: str) -> list[dict]:
    p = os.path.join(app_dir, "suspended_accounts.txt")
    entries = []
    if not os.path.exists(p):
        return entries
    try:
        with open(p, "r", encoding="utf-8") as f:
            for idx, line in enumerate(f):
                raw = line.strip()
                if not raw:
                    continue
                parts = raw.split(",", 2)
                handle = parts[0].strip()
                ts = parts[1].strip() if len(parts) > 1 else ""
                reason = parts[2].strip() if len(parts) > 2 else ""
                entries.append({
                    "index": idx,
                    "handle": handle,
                    "clean_id": handle.replace("@", "").strip().lower(),
                    "timestamp": ts,
                    "reason": reason,
                    "raw": raw,
                })
    except Exception:
        pass
    return entries


def read_json_safe(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


class XBotControlPanel:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.app_dir = get_app_dir()
        os.chdir(self.app_dir)
        self.sessions_root = os.path.join(self.app_dir, "browser_sessions")
        self.log_root = os.path.join(self.app_dir, "log")
        self.ss_root = os.path.join(self.app_dir, "screenshots")
        self.images_root = os.path.join(self.app_dir, "images")
        for d in (self.sessions_root, self.log_root, self.ss_root, self.images_root):
            os.makedirs(d, exist_ok=True)

        self.bot_proc: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.external_running = False
        self._preview_photo_ref = None
        self._img_preview_ref = None
        self._theme_tweets_cache: list[dict] = []
        self._selected_post_index: int | None = None

        self.root.title("X Auto Tweet Bot コントロールパネル (統合GUI)")
        self.root.geometry("1160x780")
        self.root.minsize(920, 620)
        self.root.configure(bg="#0f172a")

        self._setup_styles()
        self._build_ui()

        self.refresh_all_views()

        self.root.after(150, self._poll_log_queue)
        self.root.after(1500, self._poll_bot_status)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        default_font = ("Yu Gothic UI", 9)
        style.configure(".", font=default_font)
        style.configure("TNotebook", background="#0f172a", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background="#1e293b",
            foreground="#cbd5e1",
            padding=[12, 7],
            font=("Yu Gothic UI", 9, "bold"),
        )
        style.map("TNotebook.Tab", background=[("selected", "#1e3a8a")], foreground=[("selected", "#ffffff")])
        style.configure("Dark.TFrame", background="#0f172a")
        style.configure(
            "Treeview",
            background="#0b1220",
            foreground="#e2e8f0",
            fieldbackground="#0b1220",
            rowheight=25,
            font=("Yu Gothic UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#1e3a8a",
            foreground="#ffffff",
            font=("Yu Gothic UI", 9, "bold"),
        )
        style.map("Treeview", background=[("selected", "#2563eb")])

    def _build_ui(self):
        # ===== Top Control Bar =====
        header = tk.Frame(self.root, bg="#111827", pady=10, padx=14, highlightthickness=1, highlightbackground="#1f2937")
        header.pack(fill=tk.X)

        left_hdr = tk.Frame(header, bg="#111827")
        left_hdr.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            left_hdr,
            text="X Auto Tweet Bot コントロールパネル",
            font=("Yu Gothic UI", 13, "bold"),
            fg="#ffffff",
            bg="#111827",
        ).pack(anchor="w")

        info_row = tk.Frame(left_hdr, bg="#111827")
        info_row.pack(anchor="w", pady=(3, 0))

        self.status_badge = tk.Label(
            info_row,
            text="● 停止中",
            font=("Yu Gothic UI", 9, "bold"),
            fg="#cbd5e1",
            bg="#334155",
            padx=10,
            pady=2,
        )
        self.status_badge.pack(side=tk.LEFT)

        self.active_task_lbl = tk.Label(
            info_row,
            text="",
            font=("Yu Gothic UI", 9, "bold"),
            fg="#38bdf8",
            bg="#111827",
            padx=8,
        )
        self.active_task_lbl.pack(side=tk.LEFT)

        tk.Label(
            info_row,
            text=f"フォルダ: {self.app_dir}",
            font=("Yu Gothic UI", 8),
            fg="#94a3b8",
            bg="#111827",
        ).pack(side=tk.LEFT)

        btn_bar = tk.Frame(header, bg="#111827")
        btn_bar.pack(side=tk.RIGHT)

        self.btn_start = tk.Button(
            btn_bar,
            text="▶ ボット開始",
            command=self.start_bot,
            font=("Yu Gothic UI", 10, "bold"),
            bg="#10b981",
            fg="#052e16",
            activebackground="#34d399",
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_stop = tk.Button(
            btn_bar,
            text="■ ボット停止",
            command=self.stop_bot,
            font=("Yu Gothic UI", 10, "bold"),
            bg="#ef4444",
            fg="#ffffff",
            activebackground="#f87171",
            relief=tk.FLAT,
            padx=14,
            pady=6,
            cursor="hand2",
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btn_bar,
            text="📁 フォルダ",
            command=lambda: self.open_path(self.app_dir),
            font=("Yu Gothic UI", 9, "bold"),
            bg="#334155",
            fg="#e2e8f0",
            relief=tk.FLAT,
            padx=10,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btn_bar,
            text="🗕 最小化",
            command=self.root.iconify,
            font=("Yu Gothic UI", 9),
            bg="#1e293b",
            fg="#cbd5e1",
            relief=tk.FLAT,
            padx=8,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT)

        # ===== Main Tabs =====
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_dashboard = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_themes = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_profiles = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_accounts = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_images = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_logs = ttk.Frame(self.notebook, style="Dark.TFrame")

        self.notebook.add(self.tab_dashboard, text=" 📊 稼働状況・テスト ")
        self.notebook.add(self.tab_themes, text=" 📝 テーマ・ポスト編集 ")
        self.notebook.add(self.tab_profiles, text=" 👤 Chromeプロファイル ")
        self.notebook.add(self.tab_accounts, text=" 🔑 アカウント情報 ")
        self.notebook.add(self.tab_images, text=" 🖼️ 画像＆クールダウン ")
        self.notebook.add(self.tab_logs, text=" 🖥️ ログ＆📸スクショ ")

        self._build_dashboard_tab()
        self._build_themes_tab()
        self._build_profiles_tab()
        self._build_accounts_tab()
        self._build_images_tab()
        self._build_logs_tab()

    def refresh_all_views(self):
        self.refresh_profiles()
        self.refresh_accounts()
        self.refresh_excel_files()
        self.refresh_dashboard()
        self.refresh_theme_account_list()
        self.refresh_image_folders()
        self.refresh_log_file_list()
        self.refresh_screenshot_accounts()

    # --------------------------------------------------
    # Tab 1: 稼働状況・テスト (Dashboard)
    # --------------------------------------------------
    def _build_dashboard_tab(self):
        container = tk.Frame(self.tab_dashboard, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        top_bar = tk.Frame(container, bg="#1e293b", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        top_bar.pack(fill=tk.X, pady=(0, 8))

        tk.Button(
            top_bar, text="↻ 状況更新", command=self.refresh_dashboard,
            font=("Yu Gothic UI", 9, "bold"), bg="#334155", fg="#f8fafc", relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            top_bar, text="⏸ 一時停止 / 再開 切替", command=self.toggle_pause_selected_dashboard,
            font=("Yu Gothic UI", 9, "bold"), bg="#d97706", fg="#ffffff", relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            top_bar, text="⚡ 選択を1件テスト投稿", command=self.test_post_selected_dashboard,
            font=("Yu Gothic UI", 9, "bold"), bg="#0284c7", fg="#ffffff", relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            top_bar, text="🔍 ログイン状態一括確認", command=self.run_session_check_ui,
            font=("Yu Gothic UI", 9, "bold"), bg="#7c3aed", fg="#ffffff", relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            top_bar, text="🚫 除外リスト管理", command=self.open_blacklist_modal,
            font=("Yu Gothic UI", 9, "bold"), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=10, pady=5, cursor="hand2"
        ).pack(side=tk.RIGHT)

        # Simulator sub-bar
        sim_bar = tk.Frame(container, bg="#1e293b", padx=10, pady=6, highlightthickness=1, highlightbackground="#334155")
        sim_bar.pack(fill=tk.X, pady=(0, 8))

        tk.Label(sim_bar, text="🧪 時刻シミュレーション (HH:MM 例 20:30):", fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT)
        self.sim_time_var = tk.StringVar(value=datetime.now().strftime("%H:%M"))
        tk.Entry(
            sim_bar, textvariable=self.sim_time_var, width=8, font=("Consolas", 10),
            bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569"
        ).pack(side=tk.LEFT, padx=(6, 6), ipady=2)
        tk.Button(
            sim_bar, text="この時刻で投稿対象を判定", command=self.simulate_time_schedule,
            font=("Yu Gothic UI", 9, "bold"), bg="#0f766e", fg="#ffffff", relief=tk.FLAT, padx=10, pady=3, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 12))

        self.dash_summary_lbl = tk.Label(sim_bar, text="", fg="#94a3b8", bg="#1e293b", font=("Yu Gothic UI", 9))
        self.dash_summary_lbl.pack(side=tk.RIGHT)

        # Main Dashboard Treeview
        table_card = tk.Frame(container, bg="#1e293b", padx=8, pady=8, highlightthickness=1, highlightbackground="#334155")
        table_card.pack(fill=tk.BOTH, expand=True)

        cols = ("acc", "acc_state", "theme", "rule", "remain", "out_time", "now_judge", "last_post", "next_eta")
        self.dash_tree = ttk.Treeview(table_card, columns=cols, show="headings")
        headers = [
            ("acc", "アカウント", 120),
            ("acc_state", "状態", 90),
            ("theme", "テーマ(シート)", 130),
            ("rule", "時間 / 曜日 / 間隔", 210),
            ("remain", "未投稿 / 全件", 95),
            ("out_time", "時間外", 65),
            ("now_judge", "現在判定", 130),
            ("last_post", "最終投稿", 135),
            ("next_eta", "次回予定(残り)", 125),
        ]
        for cid, txt, w in headers:
            self.dash_tree.heading(cid, text=txt)
            self.dash_tree.column(cid, width=w, anchor=tk.CENTER if cid not in ("acc", "theme", "rule") else tk.W)

        yscroll = ttk.Scrollbar(table_card, orient=tk.VERTICAL, command=self.dash_tree.yview)
        self.dash_tree.configure(yscrollcommand=yscroll.set)
        self.dash_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.dash_tree.bind("<Double-Button-1>", lambda _e: self._jump_from_dashboard_to_theme())

    def _find_account_excel_path(self, clean_id: str) -> str | None:
        for cand in (f"{clean_id}.xlsx", f"@{clean_id}.xlsx", f"tweets_{clean_id}.xlsx"):
            p = os.path.join(self.app_dir, cand)
            if os.path.exists(p):
                return p
        return None

    def _read_excel_theme_summary(self, excel_path: str, now_dt: datetime) -> list[dict]:
        summaries = []
        if not HAS_OPENPYXL or not excel_path or not os.path.exists(excel_path):
            return summaries
        try:
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            for sname in wb.sheetnames:
                if sname in LEGACY_IGNORED_SHEETS:
                    continue
                ws = wb[sname]
                a1 = normalize_text(ws.cell(1, 1).value)
                b1 = normalize_text(ws.cell(1, 2).value)
                has_banner = ("実行時間" in a1) and ("曜日" in b1 or "間隔" in normalize_text(ws.cell(1, 3).value))
                time_range = normalize_text(ws.cell(2, 1).value) or "0:00-0:00" if has_banner else "0:00-0:00"
                weekdays = normalize_text(ws.cell(2, 2).value) or "毎日" if has_banner else "毎日"
                interval = normalize_text(ws.cell(2, 3).value) or "120~360分" if has_banner else "120~360分"
                act_raw = normalize_text(ws.cell(2, 5).value).lower() if has_banner else "有効"
                enabled = act_raw not in ("無効", "停止", "off", "false", "0", "disable", "disabled", "×", "✕")

                hdr_row = 4 if has_banner else 1
                for r in range(3 if has_banner else 1, min(15, max(ws.max_row, 4)) + 1):
                    vals = [normalize_text(ws.cell(r, c).value).lower() for c in range(1, min(ws.max_column, 12) + 1)]
                    if any(v in ("ツイート文章", "ポスト文章", "ツイート本文", "本文", "text") for v in vals):
                        hdr_row = r
                        break

                text_col, pt_col, st_col, dt_col = 2, (3 if has_banner else None), (5 if has_banner else 4), (6 if has_banner else 5)
                for c in range(1, ws.max_column + 1):
                    h = normalize_text(ws.cell(hdr_row, c).value).lower()
                    if h in ("ツイート文章", "ポスト文章", "ツイート本文", "本文", "text"):
                        text_col = c
                    elif h in ("投稿時間", "投稿時間帯", "時間指定", "post_time", "time"):
                        pt_col = c
                    elif h in ("ステータス", "status"):
                        st_col = c
                    elif h in ("投稿日時", "posted_at"):
                        dt_col = c

                total, unposted, eligible, skipped_time = 0, 0, 0, 0
                last_post = ""
                for r in range(hdr_row + 1, ws.max_row + 1):
                    txt = str(ws.cell(r, text_col).value or "").strip()
                    if not txt:
                        continue
                    total += 1
                    st_val = normalize_text(ws.cell(r, st_col).value).lower()
                    dt_val = str(ws.cell(r, dt_col).value or "").strip()
                    if dt_val and dt_val > last_post:
                        last_post = dt_val
                    if st_val in ("posted", "投稿済み", "投稿済", "完了", "skipped", "スキップ"):
                        continue
                    unposted += 1
                    pt_val = normalize_text(ws.cell(r, pt_col).value) if pt_col else ""
                    if pt_val and not is_time_in_ranges(pt_val, now_dt=now_dt):
                        skipped_time += 1
                    else:
                        eligible += 1

                summaries.append({
                    "sheet_name": sname,
                    "time_range": time_range,
                    "weekdays": weekdays,
                    "interval": interval,
                    "enabled": enabled,
                    "total": total,
                    "unposted": unposted,
                    "eligible": eligible,
                    "skipped_time": skipped_time,
                    "last_post": last_post,
                })
        except Exception:
            pass
        return summaries

    def refresh_dashboard(self, sim_dt: datetime | None = None):
        if not hasattr(self, "dash_tree"):
            return
        for item in self.dash_tree.get_children():
            self.dash_tree.delete(item)
        if not HAS_OPENPYXL:
            return

        now_dt = sim_dt or datetime.now()
        paused = load_paused_set(self.app_dir)
        suspended_map = {e["clean_id"]: e for e in load_suspended_entries(self.app_dir)}
        status_json = read_json_safe(os.path.join(self.app_dir, "bot_status.json"), {})
        next_map = status_json.get("next_post_times", {}) if isinstance(status_json, dict) else {}

        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        rows_added = 0
        active_cnt, paused_cnt = 0, 0
        if os.path.exists(acc_path):
            try:
                wb = openpyxl.load_workbook(acc_path, data_only=True)
                ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
                for r in range(2, ws.max_row + 1):
                    uid = str(ws.cell(r, 1).value or "").strip()
                    if not uid:
                        continue
                    clean_id = sanitize_profile_name(uid).lower()
                    if clean_id in suspended_map:
                        acc_state = "🚫 凍結除外"
                    elif clean_id in paused:
                        acc_state = "⏸ 一時休止"
                        paused_cnt += 1
                    else:
                        acc_state = "🟢 稼働対象"
                        active_cnt += 1

                    ex_path = self._find_account_excel_path(uid.replace("@", "").strip())
                    if not ex_path:
                        self.dash_tree.insert("", tk.END, values=(clean_id, acc_state, "(Excelなし)", "-", "-", "-", "⚠️ ファイルなし", "-", "-"))
                        rows_added += 1
                        continue

                    themes = self._read_excel_theme_summary(ex_path, now_dt)
                    if not themes:
                        self.dash_tree.insert("", tk.END, values=(clean_id, acc_state, "(シートなし)", "-", "-", "-", "⚠️ 空", "-", "-"))
                        rows_added += 1
                        continue

                    for t in themes:
                        sname = t["sheet_name"]
                        if clean_id in suspended_map:
                            judge = "🚫 除外"
                        elif clean_id in paused:
                            judge = "⏸ 休止中"
                        elif not t["enabled"]:
                            judge = "⛔ テーマ無効"
                        elif not is_today_allowed(t["weekdays"], now_dt=now_dt):
                            judge = "📅 曜日外"
                        elif not is_time_in_ranges(t["time_range"], now_dt=now_dt):
                            judge = "⏳ 時間外"
                        elif t["total"] == 0:
                            judge = "📝 ポスト空"
                        elif t["eligible"] == 0 and t["skipped_time"] > 0:
                            judge = "⏳ 全件時間外"
                        else:
                            judge = f"🟢 可 ({t['eligible']}件)"

                        eta_str = "-"
                        key = f"{clean_id}|{sname}"
                        if key in next_map:
                            try:
                                target_dt = datetime.strptime(next_map[key], "%Y-%m-%d %H:%M:%S")
                                diff = int((target_dt - datetime.now()).total_seconds())
                                if diff > 0:
                                    eta_str = f"{target_dt.strftime('%H:%M')} (あと{diff//60}分)"
                                else:
                                    eta_str = "順番待ち"
                            except Exception:
                                eta_str = next_map[key]

                        rule_str = f"{t['time_range']} / {t['weekdays']} / {t['interval']}"
                        remain_str = f"{t['unposted']} / {t['total']}"
                        self.dash_tree.insert(
                            "", tk.END,
                            values=(clean_id, acc_state, sname, rule_str, remain_str, str(t["skipped_time"]), judge, t["last_post"] or "-", eta_str),
                        )
                        rows_added += 1
            except Exception:
                pass

        self.dash_summary_lbl.config(
            text=f"稼働対象: {active_cnt}件 / 休止中: {paused_cnt}件 / 表示行数: {rows_added}行 ({now_dt.strftime('%H:%M')}基準)"
        )

    def _refresh_dashboard_countdowns_only(self):
        status_json = read_json_safe(os.path.join(self.app_dir, "bot_status.json"), {})
        if not isinstance(status_json, dict):
            return
        cur_acc = status_json.get("current_account", "")
        cur_thm = status_json.get("current_theme", "")
        if cur_acc:
            self.active_task_lbl.config(text=f"▶ 処理中: @{cur_acc} ({cur_thm})")
        else:
            self.active_task_lbl.config(text="")

        next_map = status_json.get("next_post_times", {})
        if not next_map or not hasattr(self, "dash_tree"):
            return
        now = datetime.now()
        for iid in self.dash_tree.get_children():
            vals = list(self.dash_tree.item(iid, "values"))
            if len(vals) >= 9:
                key = f"{vals[0]}|{vals[2]}"
                if key in next_map:
                    try:
                        target_dt = datetime.strptime(next_map[key], "%Y-%m-%d %H:%M:%S")
                        diff = int((target_dt - now).total_seconds())
                        vals[8] = f"{target_dt.strftime('%H:%M')} (あと{diff//60}分{diff%60:02d}秒)" if diff > 0 else "順番待ち"
                        self.dash_tree.item(iid, values=vals)
                    except Exception:
                        pass

    def toggle_pause_selected_dashboard(self):
        sel = self.dash_tree.selection()
        if not sel:
            messagebox.showinfo("選択確認", "一時停止 / 再開を切り替えるアカウントの行を選択してください。")
            return
        vals = self.dash_tree.item(sel[0], "values")
        clean_id = str(vals[0]).lower()
        paused = load_paused_set(self.app_dir)
        if clean_id in paused:
            paused.remove(clean_id)
            state_msg = "稼働対象（再開）"
        else:
            paused.add(clean_id)
            state_msg = "一時休止"
        save_paused_set(self.app_dir, paused)
        self.refresh_dashboard()
        self.refresh_accounts()
        self.append_log(f"⏸ アカウント [@{clean_id}] を「{state_msg}」に変更しました。\n")

    def test_post_selected_dashboard(self):
        sel = self.dash_tree.selection()
        if not sel:
            messagebox.showinfo("選択確認", "テスト投稿したいアカウントの行を選択してください。")
            return
        vals = self.dash_tree.item(sel[0], "values")
        clean_id = str(vals[0]).strip()
        theme_name = str(vals[2]).strip()
        if theme_name.startswith("("):
            theme_name = ""

        if (self.bot_proc is not None and self.bot_proc.poll() is None) or check_any_bot_running_fast(self.app_dir):
            messagebox.showwarning("稼働中", "通常ボットが稼働中のため、停止してから1件テスト投稿を実行してください。")
            return

        if not messagebox.askyesno(
            "1件テスト投稿の確認",
            f"アカウント「@{clean_id}」" + (f"（テーマ: {theme_name}）" if theme_name else "") + " の未投稿ツイートを1件だけテスト投稿しますか？",
        ):
            return

        py_script = os.path.join(self.app_dir, "auto_tweet_playwright.py")
        exe_path = os.path.join(self.app_dir, "X_AutoTweetBot.exe")
        if os.path.isfile(py_script):
            cmd = [find_python_cli(), "-u", py_script, "--test-account", clean_id]
        elif os.path.isfile(exe_path):
            cmd = [exe_path, "--test-account", clean_id]
        else:
            messagebox.showerror("エラー", "auto_tweet_playwright.py が見つかりません。")
            return
        if theme_name:
            cmd.extend(["--test-theme", theme_name])

        self._launch_bot_subprocess(cmd, f"\n=== ⚡ 1件テスト投稿開始 (@{clean_id}) ===\n")

    def run_session_check_ui(self):
        if (self.bot_proc is not None and self.bot_proc.poll() is None) or check_any_bot_running_fast(self.app_dir):
            messagebox.showwarning("稼働中", "ボット稼働中はプロファイルが競合するため、停止してからチェックを実行してください。")
            return
        py_script = os.path.join(self.app_dir, "auto_tweet_playwright.py")
        exe_path = os.path.join(self.app_dir, "X_AutoTweetBot.exe")
        if os.path.isfile(py_script):
            cmd = [find_python_cli(), "-u", py_script, "--check-sessions"]
        elif os.path.isfile(exe_path):
            cmd = [exe_path, "--check-sessions"]
        else:
            messagebox.showerror("エラー", "auto_tweet_playwright.py が見つかりません。")
            return
        self._launch_bot_subprocess(cmd, "\n=== 🔍 全プロファイル ログイン状態チェック開始 ===\n", on_done=self.refresh_profiles)

    def simulate_time_schedule(self):
        raw = self.sim_time_var.get().strip()
        m = parse_single_time_to_minutes(raw)
        if m is None or m >= 24 * 60:
            messagebox.showwarning("時刻入力エラー", "時刻は 20:30 や 7:00 の形式で入力してください。")
            return
        now = datetime.now()
        sim_dt = now.replace(hour=m // 60, minute=m % 60, second=0, microsecond=0)
        self.refresh_dashboard(sim_dt=sim_dt)
        messagebox.showinfo("シミュレーション反映", f"{sim_dt.strftime('%H:%M')} 時点での各テーマ・ポスト判定を一覧に反映しました。")

    def _jump_from_dashboard_to_theme(self):
        sel = self.dash_tree.selection()
        if not sel:
            return
        vals = self.dash_tree.item(sel[0], "values")
        clean_id = str(vals[0]).strip()
        theme_name = str(vals[2]).strip()
        ex_path = self._find_account_excel_path(clean_id)
        if not ex_path:
            return
        fname = os.path.basename(ex_path)
        self.refresh_theme_account_list()
        if fname in self.theme_acc_combo["values"]:
            self.theme_acc_var.set(fname)
            self.on_theme_account_changed()
            if theme_name in self.theme_sheet_combo["values"]:
                self.theme_sheet_var.set(theme_name)
                self.on_theme_sheet_changed()
            self.notebook.select(self.tab_themes)

    def open_blacklist_modal(self):
        win = tk.Toplevel(self.root)
        win.title("凍結・除外アカウント管理 (suspended_accounts.txt)")
        win.geometry("680x400")
        win.configure(bg="#0f172a")
        win.transient(self.root)

        tk.Label(
            win, text="現在自動除外されているアカウント一覧（誤検知だった場合は「除外解除」できます）",
            fg="#e2e8f0", bg="#0f172a", font=("Yu Gothic UI", 9, "bold"), padx=12, pady=8
        ).pack(anchor="w")

        tree = ttk.Treeview(win, columns=("handle", "ts", "reason"), show="headings", height=10)
        tree.heading("handle", text="アカウント")
        tree.heading("ts", text="検知日時")
        tree.heading("reason", text="理由")
        tree.column("handle", width=150)
        tree.column("ts", width=160)
        tree.column("reason", width=320)
        tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        def _reload():
            for iid in tree.get_children():
                tree.delete(iid)
            for e in load_suspended_entries(self.app_dir):
                tree.insert("", tk.END, values=(e["handle"], e["timestamp"], e["reason"]))

        def _unblock():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("選択", "解除するアカウントを選択してください。", parent=win)
                return
            target_handle = str(tree.item(sel[0], "values")[0]).strip()
            target_clean = target_handle.replace("@", "").lower()
            p = os.path.join(self.app_dir, "suspended_accounts.txt")
            if os.path.exists(p):
                with open(p, "r", encoding="utf-8") as f:
                    lines = [ln for ln in f if ln.strip() and ln.split(",")[0].strip().replace("@", "").lower() != target_clean]
                with open(p, "w", encoding="utf-8") as f:
                    f.writelines(lines)
            _reload()
            self.refresh_dashboard()
            self.append_log(f"🔓 [{target_handle}] を除外リストから解除しました。（アカウント情報.xlsx に未登録の場合は再追加してください）\n")

        _reload()
        btn_f = tk.Frame(win, bg="#0f172a", pady=8, padx=12)
        btn_f.pack(fill=tk.X)
        tk.Button(btn_f, text="🔓 選択アカウントを除外解除", command=_unblock, bg="#059669", fg="#ffffff", font=("Yu Gothic UI", 9, "bold"), relief=tk.FLAT, padx=12, pady=5, cursor="hand2").pack(side=tk.LEFT)
        tk.Button(btn_f, text="閉じる", command=win.destroy, bg="#334155", fg="#e2e8f0", font=("Yu Gothic UI", 9), relief=tk.FLAT, padx=12, pady=5, cursor="hand2").pack(side=tk.RIGHT)

    # --------------------------------------------------
    # Tab 2: テーマ・ポスト編集 (アカウント×テーマ別直編集)
    # --------------------------------------------------
    def _build_themes_tab(self):
        container = tk.Frame(self.tab_themes, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        # Row 0: Account & Theme Sheet selector
        sel_bar = tk.Frame(container, bg="#1e293b", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        sel_bar.pack(fill=tk.X, pady=(0, 8))

        tk.Label(sel_bar, text="対象アカウントExcel:", fg="#f8fafc", bg="#1e293b", font=("Yu Gothic UI", 9, "bold")).pack(side=tk.LEFT)
        self.theme_acc_var = tk.StringVar()
        self.theme_acc_combo = ttk.Combobox(sel_bar, textvariable=self.theme_acc_var, state="readonly", width=22)
        self.theme_acc_combo.pack(side=tk.LEFT, padx=(6, 6))
        self.theme_acc_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_theme_account_changed())

        tk.Button(
            sel_bar, text="＋ 新規アカウントExcel作成", command=self.create_new_account_excel,
            font=("Yu Gothic UI", 8, "bold"), bg="#0284c7", fg="#ffffff", relief=tk.FLAT, padx=8, pady=4, cursor="hand2"
        ).pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(sel_bar, text="テーマ(シート):", fg="#f8fafc", bg="#1e293b", font=("Yu Gothic UI", 9, "bold")).pack(side=tk.LEFT)
        self.theme_sheet_var = tk.StringVar()
        self.theme_sheet_combo = ttk.Combobox(sel_bar, textvariable=self.theme_sheet_var, state="readonly", width=18)
        self.theme_sheet_combo.pack(side=tk.LEFT, padx=(6, 6))
        self.theme_sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_theme_sheet_changed())

        tk.Button(sel_bar, text="＋ テーマ追加", command=self.add_theme_sheet, font=("Yu Gothic UI", 8, "bold"), bg="#059669", fg="#ffffff", relief=tk.FLAT, padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(sel_bar, text="📋 複製", command=self.duplicate_theme_sheet, font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(sel_bar, text="🗑️ テーマ削除", command=self.delete_theme_sheet, font=("Yu Gothic UI", 8), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=8, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(sel_bar, text="📊 Excelを開く", command=self.open_current_theme_excel, font=("Yu Gothic UI", 8), bg="#065f46", fg="#d1fae5", relief=tk.FLAT, padx=8, pady=4, cursor="hand2").pack(side=tk.RIGHT)

        # Row 1: Selected Theme Settings Editor
        st_card = tk.Frame(container, bg="#1e293b", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        st_card.pack(fill=tk.X, pady=(0, 8))

        st_top = tk.Frame(st_card, bg="#1e293b")
        st_top.pack(fill=tk.X, pady=(0, 6))
        tk.Label(st_top, text="選択中テーマの設定 (1〜2行目)", fg="#38bdf8", bg="#1e293b", font=("Yu Gothic UI", 9, "bold")).pack(side=tk.LEFT)
        tk.Button(st_top, text="🔄 このExcel全テーマを一括リセット", command=lambda: self.reset_theme_status(all_sheets=True), font=("Yu Gothic UI", 8), bg="#475569", fg="#f8fafc", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(st_top, text="🔄 このテーマのみステータス一括リセット", command=lambda: self.reset_theme_status(all_sheets=False), font=("Yu Gothic UI", 8, "bold"), bg="#d97706", fg="#ffffff", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(st_top, text="💾 このテーマ設定を保存", command=self.save_current_theme_settings, font=("Yu Gothic UI", 9, "bold"), bg="#2563eb", fg="#ffffff", relief=tk.FLAT, padx=12, pady=3, cursor="hand2").pack(side=tk.RIGHT)

        st_fields = tk.Frame(st_card, bg="#1e293b")
        st_fields.pack(fill=tk.X)

        self.th_time_var = tk.StringVar()
        self.th_week_var = tk.StringVar()
        self.th_iv_var = tk.StringVar()
        self.th_mode_var = tk.StringVar()
        self.th_active_var = tk.StringVar()
        self.th_memo_var = tk.StringVar()

        for lbl, var, w in [
            ("実行時間:", self.th_time_var, 13),
            ("曜日:", self.th_week_var, 10),
            ("ポスト間隔:", self.th_iv_var, 12),
        ]:
            tk.Label(st_fields, text=lbl, fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
            tk.Entry(st_fields, textvariable=var, width=w, bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, padx=(0, 10), ipady=2)

        tk.Label(st_fields, text="モード:", fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Combobox(st_fields, textvariable=self.th_mode_var, values=["1: ランダム", "0: 上から順"], state="readonly", width=13).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(st_fields, text="状態:", fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Combobox(st_fields, textvariable=self.th_active_var, values=["有効", "無効"], state="readonly", width=7).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(st_fields, text="メモ:", fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
        tk.Entry(st_fields, textvariable=self.th_memo_var, width=26, bg="#0b1220", fg="#cbd5e1", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)

        # Row 2: Tweets Treeview + Inline Editor
        post_card = tk.Frame(container, bg="#1e293b", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        post_card.pack(fill=tk.BOTH, expand=True)

        p_cols = ("no", "text", "ptime", "img", "status", "dt")
        tree_wrap = tk.Frame(post_card, bg="#1e293b")
        tree_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.post_tree = ttk.Treeview(tree_wrap, columns=p_cols, show="headings", height=8)
        for cid, lbl, w in [
            ("no", "No", 45),
            ("text", "ツイート文章", 440),
            ("ptime", "投稿時間", 110),
            ("img", "画像パス", 170),
            ("status", "ステータス", 90),
            ("dt", "投稿日時", 145),
        ]:
            self.post_tree.heading(cid, text=lbl)
            self.post_tree.column(cid, width=w, anchor=tk.CENTER if cid in ("no", "ptime", "status", "dt") else tk.W)
        py_scr = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.post_tree.yview)
        self.post_tree.configure(yscrollcommand=py_scr.set)
        self.post_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        py_scr.pack(side=tk.RIGHT, fill=tk.Y)
        self.post_tree.bind("<<TreeviewSelect>>", self.on_select_post_row)

        # Inline Post Form
        edit_box = tk.Frame(post_card, bg="#0f172a", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        edit_box.pack(fill=tk.X)

        left_edit = tk.Frame(edit_box, bg="#0f172a")
        left_edit.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        tk.Label(left_edit, text="ツイート文章（改行可・スピンタックス {A|B} 対応）:", fg="#cbd5e1", bg="#0f172a", font=("Yu Gothic UI", 8, "bold")).pack(anchor="w")
        self.post_text_widget = tk.Text(
            left_edit, height=3, font=("Yu Gothic UI", 9),
            bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569"
        )
        self.post_text_widget.pack(fill=tk.X, pady=(2, 0))

        right_edit = tk.Frame(edit_box, bg="#0f172a")
        right_edit.pack(side=tk.RIGHT, fill=tk.Y)

        r1 = tk.Frame(right_edit, bg="#0f172a")
        r1.pack(fill=tk.X, pady=(0, 4))
        self.post_ptime_var = tk.StringVar()
        self.post_status_var = tk.StringVar(value="pending")
        self.post_img_var = tk.StringVar()

        tk.Label(r1, text="投稿時間(例 7～10 / 空=随時):", fg="#cbd5e1", bg="#0f172a", font=("Yu Gothic UI", 8)).pack(side=tk.LEFT)
        tk.Entry(r1, textvariable=self.post_ptime_var, width=16, bg="#0b1220", fg="#fbbf24", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, padx=(4, 10), ipady=2)
        tk.Label(r1, text="ステータス:", fg="#cbd5e1", bg="#0f172a", font=("Yu Gothic UI", 8)).pack(side=tk.LEFT)
        ttk.Combobox(r1, textvariable=self.post_status_var, values=["pending", "posted", "skipped"], state="readonly", width=10).pack(side=tk.LEFT, padx=(4, 0))

        r2 = tk.Frame(right_edit, bg="#0f172a")
        r2.pack(fill=tk.X, pady=(0, 6))
        tk.Label(r2, text="画像パス:", fg="#cbd5e1", bg="#0f172a", font=("Yu Gothic UI", 8)).pack(side=tk.LEFT)
        tk.Entry(r2, textvariable=self.post_img_var, width=28, bg="#0b1220", fg="#e2e8f0", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, padx=(4, 4), ipady=2)
        tk.Button(r2, text="📄画像", command=lambda: self._browse_post_image(folder=False), bg="#334155", fg="#e2e8f0", font=("Yu Gothic UI", 8), relief=tk.FLAT, padx=6, pady=2, cursor="hand2").pack(side=tk.LEFT, padx=(0, 3))
        tk.Button(r2, text="📁フォルダ", command=lambda: self._browse_post_image(folder=True), bg="#334155", fg="#e2e8f0", font=("Yu Gothic UI", 8), relief=tk.FLAT, padx=6, pady=2, cursor="hand2").pack(side=tk.LEFT)

        r3 = tk.Frame(right_edit, bg="#0f172a")
        r3.pack(fill=tk.X)
        tk.Button(r3, text="＋ 新規ポスト追加", command=self.add_post_row, bg="#10b981", fg="#052e16", font=("Yu Gothic UI", 9, "bold"), relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r3, text="💾 選択行を更新", command=self.update_post_row, bg="#2563eb", fg="#ffffff", font=("Yu Gothic UI", 9, "bold"), relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(r3, text="🗑️ 選択行を削除", command=self.delete_post_row, bg="#7f1d1d", fg="#fecaca", font=("Yu Gothic UI", 9, "bold"), relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT)

    def _get_current_theme_excel_path(self) -> str | None:
        fname = self.theme_acc_var.get().strip()
        if not fname:
            return None
        p = os.path.join(self.app_dir, fname)
        return p if os.path.exists(p) else None

    def refresh_theme_account_list(self):
        if not hasattr(self, "theme_acc_combo"):
            return
        files = []
        for path in sorted(glob.glob(os.path.join(self.app_dir, "*.xlsx"))):
            fn = os.path.basename(path)
            if fn.startswith("~$") or "アカウント情報" in fn or "dm_targets" in fn:
                continue
            files.append(fn)
        self.theme_acc_combo["values"] = files
        cur = self.theme_acc_var.get()
        if files and cur not in files:
            self.theme_acc_var.set(files[0])
            self.on_theme_account_changed()
        elif cur in files:
            self.on_theme_account_changed()
        elif not files:
            self.theme_acc_var.set("")
            self.theme_sheet_combo["values"] = []
            self.theme_sheet_var.set("")

    def on_theme_account_changed(self):
        ex_path = self._get_current_theme_excel_path()
        if not ex_path or not HAS_OPENPYXL:
            self.theme_sheet_combo["values"] = []
            self.theme_sheet_var.set("")
            return
        try:
            wb = openpyxl.load_workbook(ex_path, data_only=True)
            sheets = [s for s in wb.sheetnames if s not in LEGACY_IGNORED_SHEETS]
            self.theme_sheet_combo["values"] = sheets
            if sheets:
                cur_s = self.theme_sheet_var.get()
                if cur_s not in sheets:
                    self.theme_sheet_var.set(sheets[0])
                self.on_theme_sheet_changed()
            else:
                self.theme_sheet_var.set("")
        except Exception as e:
            self.append_log(f"⚠️ Excelシート一覧読込エラー: {e}\n")

    def on_theme_sheet_changed(self):
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not sname or not HAS_OPENPYXL:
            return
        try:
            wb = openpyxl.load_workbook(ex_path, data_only=True)
            if sname not in wb.sheetnames:
                return
            ws = wb[sname]
            default_st = {
                "time_range": "0:00-0:00",
                "weekdays": "毎日",
                "interval": "120~360分",
                "mode": "1: ランダム",
                "active": "有効",
                "memo": theme_helper.DEFAULT_MEMO if HAS_THEME_HELPER else "",
            }
            if HAS_THEME_HELPER:
                settings, tweets = theme_helper._read_existing_sheet_data(ws, default_st)
            else:
                settings = default_st
                tweets = []

            self.th_time_var.set(settings.get("time_range", "0:00-0:00"))
            self.th_week_var.set(settings.get("weekdays", "毎日"))
            self.th_iv_var.set(settings.get("interval", "120~360分"))
            self.th_mode_var.set(settings.get("mode", "1: ランダム"))
            self.th_active_var.set(settings.get("active", "有効"))
            self.th_memo_var.set(settings.get("memo", ""))

            self._theme_tweets_cache = tweets
            self._selected_post_index = None
            for iid in self.post_tree.get_children():
                self.post_tree.delete(iid)
            for idx, t in enumerate(tweets):
                preview_txt = t["text"].replace("\n", " ⏎ ")
                if len(preview_txt) > 70:
                    preview_txt = preview_txt[:70] + "..."
                self.post_tree.insert(
                    "", tk.END, iid=str(idx),
                    values=(t.get("no") or (idx + 1), preview_txt, t.get("post_time", ""), t.get("image", ""), t.get("status") or "pending", t.get("posted_at", ""))
                )
        except Exception as e:
            messagebox.showerror("読込エラー", f"シート「{sname}」の読み込みに失敗しました:\n{e}")

    def _save_theme_sheet_back(self, settings: dict, tweets: list[dict]) -> bool:
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not sname or not HAS_OPENPYXL:
            return False
        try:
            wb = openpyxl.load_workbook(ex_path)
            if sname not in wb.sheetnames:
                wb.create_sheet(title=sname)
            ws = wb[sname]
            if HAS_THEME_HELPER:
                theme_helper._apply_theme_sheet_design(ws, settings, tweets)
            else:
                ws.cell(2, 1, settings.get("time_range", "0:00-0:00"))
                ws.cell(2, 2, settings.get("weekdays", "毎日"))
                ws.cell(2, 3, settings.get("interval", "120~360分"))
                ws.cell(2, 4, settings.get("mode", "1: ランダム"))
                ws.cell(2, 5, settings.get("active", "有効"))
            wb.save(ex_path)
            self.on_theme_sheet_changed()
            self.refresh_dashboard()
            return True
        except PermissionError:
            messagebox.showerror("保存エラー", f"「{os.path.basename(ex_path)}」が Excel で開かれています。\nExcel を閉じてから再度お試しください。")
            return False
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{e}")
            return False

    def _collect_theme_settings_from_ui(self) -> dict:
        return {
            "time_range": self.th_time_var.get().strip() or "0:00-0:00",
            "weekdays": self.th_week_var.get().strip() or "毎日",
            "interval": self.th_iv_var.get().strip() or "120~360分",
            "mode": self.th_mode_var.get().strip() or "1: ランダム",
            "active": self.th_active_var.get().strip() or "有効",
            "memo": self.th_memo_var.get().strip() or (theme_helper.DEFAULT_MEMO if HAS_THEME_HELPER else ""),
        }

    def save_current_theme_settings(self):
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not sname:
            messagebox.showwarning("選択確認", "アカウントExcelとテーマシートを選択してください。")
            return
        if self._save_theme_sheet_back(self._collect_theme_settings_from_ui(), list(self._theme_tweets_cache)):
            self.append_log(f"💾 [{os.path.basename(ex_path)} / {sname}] のテーマ設定を保存しました。\n")
            messagebox.showinfo("保存完了", f"「{os.path.basename(ex_path)}」のテーマ「{sname}」設定を保存しました。")

    def reset_theme_status(self, all_sheets: bool = False):
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not HAS_OPENPYXL:
            return
        if not all_sheets and not sname:
            return
        scope_label = f"「{os.path.basename(ex_path)}」の全テーマ" if all_sheets else f"「{os.path.basename(ex_path)} / {sname}」"
        if not messagebox.askyesno("ステータス一括リセット確認", f"{scope_label} のステータスをすべて未投稿(pending)にリセットしますか？"):
            return
        try:
            wb = openpyxl.load_workbook(ex_path)
            targets = [s for s in wb.sheetnames if s not in LEGACY_IGNORED_SHEETS] if all_sheets else [sname]
            for ts in targets:
                if ts not in wb.sheetnames:
                    continue
                ws = wb[ts]
                if HAS_THEME_HELPER:
                    st, tws = theme_helper._read_existing_sheet_data(ws, {"time_range": "0:00-0:00", "weekdays": "毎日", "interval": "120~360分", "mode": "1: ランダム", "active": "有効"})
                    for item in tws:
                        item["status"] = "pending"
                        item["posted_at"] = ""
                    theme_helper._apply_theme_sheet_design(ws, st, tws)
            wb.save(ex_path)
            self.on_theme_sheet_changed()
            self.refresh_dashboard()
            messagebox.showinfo("リセット完了", f"{scope_label} をすべて未投稿にリセットしました。")
        except PermissionError:
            messagebox.showerror("エラー", f"「{os.path.basename(ex_path)}」をExcelで閉じているか確認してください。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def create_new_account_excel(self, prefill_name: str = ""):
        if not HAS_OPENPYXL:
            return
        raw = prefill_name or simpledialog.askstring(
            "新規アカウントExcel作成",
            "作成するアカウントIDまたはファイル名（例: my_account → my_account.xlsx）:",
            parent=self.root,
        )
        if not raw:
            return
        clean = sanitize_profile_name(raw)
        if not clean:
            return
        fname = clean if clean.lower().endswith(".xlsx") else f"{clean}.xlsx"
        target_path = os.path.join(self.app_dir, fname)
        if os.path.exists(target_path):
            messagebox.showinfo("既存ファイル", f"「{fname}」はすでに存在します。")
            self.refresh_theme_account_list()
            self.theme_acc_var.set(fname)
            self.on_theme_account_changed()
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "テーマ1"
            default_st = {
                "time_range": "0:00-0:00",
                "weekdays": "毎日",
                "interval": "120~360分",
                "mode": "1: ランダム",
                "active": "有効",
                "memo": theme_helper.DEFAULT_MEMO if HAS_THEME_HELPER else "",
            }
            sample_tweets = [
                {"no": 1, "text": "サンプル投稿です（GUIやExcelから書き換えてください）", "post_time": "", "image": "", "status": "pending", "posted_at": ""},
            ]
            if HAS_THEME_HELPER:
                theme_helper._apply_theme_sheet_design(ws, default_st, sample_tweets)
            wb.save(target_path)
            self.refresh_excel_files()
            self.refresh_theme_account_list()
            self.theme_acc_var.set(fname)
            self.on_theme_account_changed()
            self.refresh_dashboard()
            self.refresh_accounts()
            messagebox.showinfo("作成完了", f"テンプレートから「{fname}」を作成しました！")
        except Exception as e:
            messagebox.showerror("作成エラー", str(e))

    def add_theme_sheet(self):
        ex_path = self._get_current_theme_excel_path()
        if not ex_path or not HAS_OPENPYXL:
            messagebox.showwarning("選択確認", "先にアカウントExcelを選択してください。")
            return
        new_name = simpledialog.askstring("新規テーマシート追加", "新しいテーマ（シート）名（例: 朝挨拶 / 商品紹介）:", parent=self.root)
        if not new_name:
            return
        new_name = re.sub(r'[\\/*?:\[\]]', "_", new_name.strip())[:31]
        if not new_name:
            return
        try:
            wb = openpyxl.load_workbook(ex_path)
            if new_name in wb.sheetnames:
                messagebox.showerror("重複", f"シート「{new_name}」はすでに存在します。")
                return
            ws = wb.create_sheet(title=new_name)
            default_st = {
                "time_range": "0:00-0:00",
                "weekdays": "毎日",
                "interval": "120~360分",
                "mode": "1: ランダム",
                "active": "有効",
                "memo": theme_helper.DEFAULT_MEMO if HAS_THEME_HELPER else "",
            }
            if HAS_THEME_HELPER:
                theme_helper._apply_theme_sheet_design(ws, default_st, [])
            wb.save(ex_path)
            self.on_theme_account_changed()
            self.theme_sheet_var.set(new_name)
            self.on_theme_sheet_changed()
            self.refresh_dashboard()
        except PermissionError:
            messagebox.showerror("エラー", "Excelを閉じてから実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def duplicate_theme_sheet(self):
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not sname or not HAS_OPENPYXL:
            return
        new_name = simpledialog.askstring("テーマ複製", f"「{sname}」の複製先シート名:", initialvalue=f"{sname}_コピー"[:31], parent=self.root)
        if not new_name:
            return
        new_name = re.sub(r'[\\/*?:\[\]]', "_", new_name.strip())[:31]
        try:
            wb = openpyxl.load_workbook(ex_path)
            if new_name in wb.sheetnames:
                messagebox.showerror("重複", f"「{new_name}」はすでに存在します。")
                return
            copied = wb.copy_worksheet(wb[sname])
            copied.title = new_name
            wb.save(ex_path)
            self.on_theme_account_changed()
            self.theme_sheet_var.set(new_name)
            self.on_theme_sheet_changed()
            self.refresh_dashboard()
        except PermissionError:
            messagebox.showerror("エラー", "Excelを閉じてから実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def delete_theme_sheet(self):
        ex_path = self._get_current_theme_excel_path()
        sname = self.theme_sheet_var.get().strip()
        if not ex_path or not sname or not HAS_OPENPYXL:
            return
        try:
            wb = openpyxl.load_workbook(ex_path)
            if len(wb.sheetnames) <= 1:
                messagebox.showwarning("削除不可", "最後の1シートは削除できません。")
                return
            if not messagebox.askyesno("削除確認", f"「{os.path.basename(ex_path)}」からテーマ「{sname}」を削除しますか？"):
                return
            wb.remove(wb[sname])
            wb.save(ex_path)
            self.on_theme_account_changed()
            self.refresh_dashboard()
        except PermissionError:
            messagebox.showerror("エラー", "Excelを閉じてから実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def open_current_theme_excel(self):
        ex_path = self._get_current_theme_excel_path()
        if ex_path:
            self.open_path(ex_path)

    def _browse_post_image(self, folder: bool = False):
        if folder:
            chosen = filedialog.askdirectory(initialdir=self.images_root, title="画像フォルダを選択")
        else:
            chosen = filedialog.askopenfilename(
                initialdir=self.images_root,
                title="添付画像を選択",
                filetypes=[("Image files", "*.png *.jpg *.jpeg *.webp *.gif"), ("All files", "*.*")],
            )
        if not chosen:
            return
        try:
            rel = os.path.relpath(chosen, self.app_dir).replace("\\", "/")
            if not rel.startswith(".."):
                if rel.startswith("images/"):
                    sub = rel[len("images/"):]
                    chosen_str = sub if folder else rel
                else:
                    chosen_str = rel
            else:
                chosen_str = chosen
        except Exception:
            chosen_str = chosen
        self.post_img_var.set(chosen_str)

    def on_select_post_row(self, _event=None):
        sel = self.post_tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
        except Exception:
            return
        if not (0 <= idx < len(self._theme_tweets_cache)):
            return
        self._selected_post_index = idx
        t = self._theme_tweets_cache[idx]
        self.post_text_widget.delete("1.0", tk.END)
        self.post_text_widget.insert("1.0", t.get("text", ""))
        self.post_ptime_var.set(t.get("post_time", ""))
        self.post_img_var.set(t.get("image", ""))
        st_val = t.get("status") or "pending"
        self.post_status_var.set(st_val if st_val in ("pending", "posted", "skipped") else "pending")

    def add_post_row(self):
        text = self.post_text_widget.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("入力確認", "ツイート文章を入力してください。")
            return
        tweets = list(self._theme_tweets_cache)
        tweets.append({
            "no": len(tweets) + 1,
            "text": text,
            "post_time": self.post_ptime_var.get().strip(),
            "image": self.post_img_var.get().strip(),
            "status": self.post_status_var.get().strip() or "pending",
            "posted_at": "",
        })
        self._save_theme_sheet_back(self._collect_theme_settings_from_ui(), tweets)

    def update_post_row(self):
        if self._selected_post_index is None or not (0 <= self._selected_post_index < len(self._theme_tweets_cache)):
            messagebox.showinfo("選択確認", "上の表から編集する行を選択してください。")
            return
        text = self.post_text_widget.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("入力確認", "ツイート文章を入力してください。")
            return
        tweets = list(self._theme_tweets_cache)
        old = tweets[self._selected_post_index]
        tweets[self._selected_post_index] = {
            "no": old.get("no") or (self._selected_post_index + 1),
            "text": text,
            "post_time": self.post_ptime_var.get().strip(),
            "image": self.post_img_var.get().strip(),
            "status": self.post_status_var.get().strip() or "pending",
            "posted_at": "" if (self.post_status_var.get().strip() == "pending") else old.get("posted_at", ""),
        }
        self._save_theme_sheet_back(self._collect_theme_settings_from_ui(), tweets)

    def delete_post_row(self):
        if self._selected_post_index is None or not (0 <= self._selected_post_index < len(self._theme_tweets_cache)):
            messagebox.showinfo("選択確認", "削除する行を選択してください。")
            return
        tweets = list(self._theme_tweets_cache)
        tweets.pop(self._selected_post_index)
        for i, item in enumerate(tweets, 1):
            item["no"] = i
        self._save_theme_sheet_back(self._collect_theme_settings_from_ui(), tweets)

    # --------------------------------------------------
    # Tab 3: Chromeプロファイル管理
    # --------------------------------------------------
    def _build_profiles_tab(self):
        container = tk.Frame(self.tab_profiles, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        hdr = tk.Frame(left, bg="#1e293b")
        hdr.pack(fill=tk.X, pady=(0, 8))
        tk.Label(hdr, text="保存済みプロファイル一覧 (browser_sessions)", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(hdr, text="↻ 更新", command=self.refresh_profiles, font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=2, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(hdr, text="🔍 ログイン状態チェック", command=self.run_session_check_ui, font=("Yu Gothic UI", 9, "bold"), bg="#7c3aed", fg="#ffffff", relief=tk.FLAT, padx=8, pady=2, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        list_frame = tk.Frame(left, bg="#1e293b")
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.profile_tree = ttk.Treeview(list_frame, columns=("name", "login", "checked"), show="headings")
        self.profile_tree.heading("name", text="プロファイル名")
        self.profile_tree.heading("login", text="Xログイン状態")
        self.profile_tree.heading("checked", text="最終チェック")
        self.profile_tree.column("name", width=180)
        self.profile_tree.column("login", width=130, anchor=tk.CENTER)
        self.profile_tree.column("checked", width=145, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.profile_tree.yview)
        self.profile_tree.configure(yscrollcommand=scrollbar.set)
        self.profile_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.profile_tree.bind("<<TreeviewSelect>>", self.on_select_profile)
        self.profile_tree.bind("<Double-Button-1>", lambda _e: self.launch_chrome_for_profile())

        btn_row_left = tk.Frame(left, bg="#1e293b")
        btn_row_left.pack(fill=tk.X, pady=(10, 0))

        tk.Button(btn_row_left, text="✏️ 名前変更", command=self.rename_selected_profile, font=("Yu Gothic UI", 9, "bold"), bg="#334155", fg="#f8fafc", relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_row_left, text="🗑️ 選択削除", command=self.delete_selected_profile, font=("Yu Gothic UI", 9, "bold"), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_row_left, text="📂 保存先フォルダ", command=lambda: self.open_path(self.sessions_root), font=("Yu Gothic UI", 9), bg="#334155", fg="#cbd5e1", relief=tk.FLAT, padx=10, pady=6, cursor="hand2").pack(side=tk.RIGHT)

        right = tk.Frame(container, bg="#1e293b", padx=16, pady=16, width=360, highlightthickness=1, highlightbackground="#334155")
        right.pack(side=tk.RIGHT, fill=tk.Y)
        right.pack_propagate(False)

        tk.Label(right, text="Chrome 手動起動（ログイン・Xプロフィール編集）", font=("Yu Gothic UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", wraplength=320, justify="left").pack(anchor="w")
        tk.Label(
            right,
            text="・一覧からプロファイルを選ぶか、新しい名前を入力してChromeを起動します。\n"
                 "・GUIやボット停止とは完全独立したプロセスで開くため、ボット停止時に閉じません。\n"
                 "・開いたChromeを「×」で閉じるとセッションが保存され、自動投稿ボットに引き継がれます。",
            font=("Yu Gothic UI", 9),
            fg="#cbd5e1",
            bg="#1e293b",
            justify="left",
            wraplength=320,
        ).pack(anchor="w", pady=(8, 14))

        tk.Label(right, text="プロファイル名（@なしID推奨 例: wqwq1320）", font=("Yu Gothic UI", 9, "bold"), fg="#e2e8f0", bg="#1e293b").pack(anchor="w")
        self.profile_name_var = tk.StringVar()
        self.entry_profile_name = tk.Entry(
            right, textvariable=self.profile_name_var, font=("Consolas", 11),
            bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569",
        )
        self.entry_profile_name.pack(fill=tk.X, ipady=6, pady=(4, 12))

        tk.Button(
            right, text="🌐 このプロファイルで Chrome を開く", command=self.launch_chrome_for_profile,
            font=("Yu Gothic UI", 10, "bold"), bg="#0284c7", fg="#ffffff", activebackground="#0ea5e9", relief=tk.FLAT, pady=10, cursor="hand2",
        ).pack(fill=tk.X, pady=(0, 10))

    # --------------------------------------------------
    # Tab 4: アカウント情報・Excel一覧
    # --------------------------------------------------
    def _build_accounts_tab(self):
        container = tk.Frame(self.tab_accounts, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        top_card = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        top_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        hdr = tk.Frame(top_card, bg="#1e293b")
        hdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(hdr, text="アカウント情報.xlsx 管理", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(hdr, text="📊 アカウント情報.xlsx をExcelで開く", command=lambda: self.open_path(os.path.join(self.app_dir, "アカウント情報.xlsx")), font=("Yu Gothic UI", 9), bg="#065f46", fg="#d1fae5", relief=tk.FLAT, padx=10, pady=3, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(hdr, text="↻ 再読込", command=self.refresh_accounts, font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        cols = ("id", "pw", "name", "token", "state", "excel")
        self.acc_tree = ttk.Treeview(top_card, columns=cols, show="headings", height=6)
        for cid, txt, w in [
            ("id", "アカウントID", 140),
            ("pw", "アカウントpw", 110),
            ("name", "アカウント名", 140),
            ("token", "auth_token", 160),
            ("state", "稼働状態", 100),
            ("excel", "対応Excel", 150),
        ]:
            self.acc_tree.heading(cid, text=txt)
            self.acc_tree.column(cid, width=w)
        self.acc_tree.pack(fill=tk.BOTH, expand=True)
        self.acc_tree.bind("<<TreeviewSelect>>", self.on_select_account)

        form = tk.Frame(top_card, bg="#1e293b", pady=8)
        form.pack(fill=tk.X)

        self.acc_id_var = tk.StringVar()
        self.acc_pw_var = tk.StringVar()
        self.acc_name_var = tk.StringVar()
        self.acc_tok_var = tk.StringVar()

        for lbl, var, w in [
            ("ID:", self.acc_id_var, 15),
            ("PW:", self.acc_pw_var, 13),
            ("表示名:", self.acc_name_var, 15),
            ("auth_token:", self.acc_tok_var, 18),
        ]:
            tk.Label(form, text=lbl, fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
            tk.Entry(form, textvariable=var, width=w, bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, padx=(0, 8), ipady=3)

        tk.Button(form, text="＋ 追加/上書き", command=self.save_account_row, font=("Yu Gothic UI", 9, "bold"), bg="#2563eb", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(form, text="📄 Excel作成", command=lambda: self.create_new_account_excel(self.acc_id_var.get().strip()), font=("Yu Gothic UI", 9, "bold"), bg="#059669", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(form, text="削除", command=self.delete_account_row, font=("Yu Gothic UI", 9, "bold"), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT)

        bot_card = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        bot_card.pack(fill=tk.BOTH, expand=True)

        bhdr = tk.Frame(bot_card, bg="#1e293b")
        bhdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(bhdr, text="投稿用 Excelファイル一覧 (テーマ別シート)", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(bhdr, text="🎨 全Excelをテーマ形式＆色付けに更新", command=self.run_excel_theme_updater, font=("Yu Gothic UI", 9, "bold"), bg="#d97706", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(bhdr, text="📊 選択したExcelを開く", command=self.open_selected_excel, font=("Yu Gothic UI", 9, "bold"), bg="#059669", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        self.excel_listbox = tk.Listbox(
            bot_card, height=5, font=("Consolas", 10), bg="#0b1220", fg="#e2e8f0",
            selectbackground="#2563eb", selectforeground="#ffffff", borderwidth=0, highlightthickness=1, highlightbackground="#334155",
        )
        self.excel_listbox.pack(fill=tk.BOTH, expand=True)
        self.excel_listbox.bind("<Double-Button-1>", lambda _e: self.open_selected_excel())

    # --------------------------------------------------
    # Tab 5: 画像＆クールダウン管理
    # --------------------------------------------------
    def _build_images_tab(self):
        container = tk.Frame(self.tab_images, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        top_bar = tk.Frame(container, bg="#1e293b", padx=10, pady=8, highlightthickness=1, highlightbackground="#334155")
        top_bar.pack(fill=tk.X, pady=(0, 8))

        tk.Button(top_bar, text="↻ 画像・履歴更新", command=self.refresh_image_folders, font=("Yu Gothic UI", 9, "bold"), bg="#334155", fg="#f8fafc", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(top_bar, text="📂 imagesフォルダを開く", command=lambda: self.open_subfolder("images"), font=("Yu Gothic UI", 9), bg="#065f46", fg="#d1fae5", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(top_bar, text="➕ 画像ファイルをコピー追加", command=self.copy_images_to_selected_folder, font=("Yu Gothic UI", 9, "bold"), bg="#0284c7", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(top_bar, text="🗑️ クールダウン履歴リセット", command=self.reset_image_history, font=("Yu Gothic UI", 9, "bold"), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT)

        body = tk.Frame(container, bg="#0f172a")
        body.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(body, bg="#1e293b", padx=10, pady=10, width=250, highlightthickness=1, highlightbackground="#334155")
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        left.pack_propagate(False)

        tk.Label(left, text="画像フォルダ (images/)", fg="#ffffff", bg="#1e293b", font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", pady=(0, 6))
        self.img_folder_listbox = tk.Listbox(
            left, font=("Consolas", 10), bg="#0b1220", fg="#f8fafc",
            selectbackground="#2563eb", selectforeground="#ffffff", borderwidth=0, highlightthickness=1, highlightbackground="#334155",
        )
        self.img_folder_listbox.pack(fill=tk.BOTH, expand=True)
        self.img_folder_listbox.bind("<<ListboxSelect>>", lambda _e: self.on_select_image_folder())

        right = tk.Frame(body, bg="#1e293b", padx=10, pady=10, highlightthickness=1, highlightbackground="#334155")
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Label(right, text="画像一覧 ＆ クールダウン状態（自垢30日 / 他垢7日）", fg="#38bdf8", bg="#1e293b", font=("Yu Gothic UI", 9, "bold")).pack(anchor="w", pady=(0, 6))

        self.img_tree = ttk.Treeview(right, columns=("fname", "count", "cd", "last"), show="headings", height=10)
        for cid, lbl, w in [
            ("fname", "画像ファイル名", 240),
            ("count", "使用回数", 75),
            ("cd", "クールダウン状況", 230),
            ("last", "最終使用", 150),
        ]:
            self.img_tree.heading(cid, text=lbl)
            self.img_tree.column(cid, width=w, anchor=tk.CENTER if cid in ("count", "last") else tk.W)
        self.img_tree.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.img_tree.bind("<Double-Button-1>", lambda _e: self.open_selected_image_file())

        prev_row = tk.Frame(right, bg="#0b1220", height=160, highlightthickness=1, highlightbackground="#334155")
        prev_row.pack(fill=tk.X)
        prev_row.pack_propagate(False)
        self.img_preview_lbl = tk.Label(prev_row, text="画像を選択するとプレビュー表示（ダブルクリックで外部ビューア）", fg="#94a3b8", bg="#0b1220", font=("Yu Gothic UI", 9))
        self.img_preview_lbl.pack(fill=tk.BOTH, expand=True)
        self.img_tree.bind("<<TreeviewSelect>>", lambda _e: self._preview_selected_image_tree())

    def refresh_image_folders(self):
        if not hasattr(self, "img_folder_listbox"):
            return
        self.img_folder_listbox.delete(0, tk.END)
        folders = ["."]
        if os.path.isdir(self.images_root):
            for entry in sorted(os.listdir(self.images_root)):
                if os.path.isdir(os.path.join(self.images_root, entry)):
                    folders.append(entry)
        for f in folders:
            self.img_folder_listbox.insert(tk.END, "(images直下)" if f == "." else f)
        if folders:
            self.img_folder_listbox.selection_set(0)
            self.on_select_image_folder()

    def _get_selected_images_dir(self) -> str:
        sel = self.img_folder_listbox.curselection()
        if not sel:
            return self.images_root
        val = self.img_folder_listbox.get(sel[0])
        return self.images_root if val == "(images直下)" else os.path.join(self.images_root, val)

    def on_select_image_folder(self):
        target_dir = self._get_selected_images_dir()
        for iid in self.img_tree.get_children():
            self.img_tree.delete(iid)
        history = read_json_safe(os.path.join(self.app_dir, "image_history.json"), {})
        if not isinstance(history, dict):
            history = {}

        exts = ("*.jpg", "*.jpeg", "*.png", "*.webp", "*.gif")
        files = []
        for ex in exts:
            files.extend(glob.glob(os.path.join(target_dir, ex)))
            files.extend(glob.glob(os.path.join(target_dir, ex.upper())))
        files = sorted(set(files))
        now = datetime.now()

        for fp in files:
            fn = os.path.basename(fp)
            recs = history.get(fn, [])
            last_str = "-"
            cd_msgs = []
            if recs:
                try:
                    latest = max(recs, key=lambda r: r.get("used_at", ""))
                    last_str = f"{latest.get('used_at', '')} (@{latest.get('account', '')})"
                    for r in recs:
                        dt = datetime.strptime(r["used_at"], "%Y-%m-%d %H:%M:%S")
                        diff_d = (now - dt).total_seconds() / 86400.0
                        if diff_d < IMAGE_COOLDOWN_SAME_DAYS:
                            cd_msgs.append(f"@{r.get('account')}({int(IMAGE_COOLDOWN_SAME_DAYS - diff_d)}日待)")
                except Exception:
                    pass
            cd_str = ", ".join(cd_msgs[:2]) if cd_msgs else "🟢 使用可"
            self.img_tree.insert("", tk.END, values=(fn, len(recs), cd_str, last_str), tags=(fp,))

    def _preview_selected_image_tree(self):
        sel = self.img_tree.selection()
        if not sel:
            return
        tags = self.img_tree.item(sel[0], "tags")
        if not tags:
            return
        self._load_image_into_label(tags[0], self.img_preview_lbl, max_w=420, max_h=145, is_img_tab=True)

    def open_selected_image_file(self):
        sel = self.img_tree.selection()
        if not sel:
            return
        tags = self.img_tree.item(sel[0], "tags")
        if tags:
            self.open_path(tags[0])

    def copy_images_to_selected_folder(self):
        target_dir = self._get_selected_images_dir()
        os.makedirs(target_dir, exist_ok=True)
        paths = filedialog.askopenfilenames(
            title="追加する画像を選択（複数可）",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.webp *.gif"), ("All files", "*.*")],
        )
        if not paths:
            return
        cnt = 0
        for src in paths:
            try:
                shutil.copy2(src, os.path.join(target_dir, os.path.basename(src)))
                cnt += 1
            except Exception:
                pass
        self.on_select_image_folder()
        messagebox.showinfo("コピー完了", f"{cnt} 件の画像を {os.path.relpath(target_dir, self.app_dir)} に追加しました。")

    def reset_image_history(self):
        p = os.path.join(self.app_dir, "image_history.json")
        if not os.path.exists(p):
            messagebox.showinfo("履歴なし", "クールダウン履歴 (image_history.json) はまだありません。")
            return
        if not messagebox.askyesno("リセット確認", "image_history.json の画像クールダウン履歴をすべてリセットしますか？"):
            return
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump({}, f)
            self.on_select_image_folder()
            messagebox.showinfo("完了", "画像クールダウン履歴をリセットしました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    # --------------------------------------------------
    # Tab 6: ログ＆スクショビューア
    # --------------------------------------------------
    def _build_logs_tab(self):
        container = tk.Frame(self.tab_logs, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        paned = tk.PanedWindow(container, orient=tk.HORIZONTAL, bg="#0f172a", sashwidth=6, sashrelief=tk.FLAT)
        paned.pack(fill=tk.BOTH, expand=True)

        # Left Pane: Per-account Log viewer + Live output
        left = tk.Frame(paned, bg="#1e293b", padx=8, pady=8, highlightthickness=1, highlightbackground="#334155")
        paned.add(left, minsize=460, stretch="always")

        lbar = tk.Frame(left, bg="#1e293b")
        lbar.pack(fill=tk.X, pady=(0, 6))

        tk.Label(lbar, text="ログ表示:", font=("Yu Gothic UI", 9, "bold"), fg="#f8fafc", bg="#1e293b").pack(side=tk.LEFT)
        self.log_source_var = tk.StringVar(value="(リアルタイム出力)")
        self.log_source_combo = ttk.Combobox(lbar, textvariable=self.log_source_var, state="readonly", width=22)
        self.log_source_combo.pack(side=tk.LEFT, padx=(4, 6))
        self.log_source_combo.bind("<<ComboboxSelected>>", lambda _e: self.reload_log_view())

        self.log_err_only_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            lbar, text="エラー/警告のみ", variable=self.log_err_only_var, command=self.reload_log_view,
            fg="#fbbf24", bg="#1e293b", selectcolor="#0b1220", activebackground="#1e293b", activeforeground="#fbbf24", font=("Yu Gothic UI", 8)
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(lbar, text="📂 logフォルダ", command=lambda: self.open_subfolder("log"), font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=7, pady=3, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(lbar, text="↻ 更新", command=lambda: [self.refresh_log_file_list(), self.reload_log_view()], font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=7, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 4))
        tk.Button(lbar, text="消去", command=lambda: self.log_text.delete("1.0", tk.END), font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=7, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 4))

        self.log_text = scrolledtext.ScrolledText(
            left, bg="#020617", fg="#e2e8f0", font=("Consolas", 9),
            insertbackground="#ffffff", borderwidth=0, highlightthickness=1, highlightbackground="#334155",
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config("err", foreground="#f87171")
        self.log_text.tag_config("warn", foreground="#fbbf24")
        self.log_text.tag_config("ok", foreground="#34d399")
        self.log_text.tag_config("info_hi", foreground="#38bdf8")

        # Right Pane: Per-account Screenshot viewer
        right = tk.Frame(paned, bg="#1e293b", padx=8, pady=8, highlightthickness=1, highlightbackground="#334155")
        paned.add(right, minsize=340)

        rbar = tk.Frame(right, bg="#1e293b")
        rbar.pack(fill=tk.X, pady=(0, 6))
        tk.Label(rbar, text="スクショ:", font=("Yu Gothic UI", 9, "bold"), fg="#f8fafc", bg="#1e293b").pack(side=tk.LEFT)
        self.ss_acc_var = tk.StringVar(value="(すべて)")
        self.ss_acc_combo = ttk.Combobox(rbar, textvariable=self.ss_acc_var, state="readonly", width=16)
        self.ss_acc_combo.pack(side=tk.LEFT, padx=(4, 4))
        self.ss_acc_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_screenshot_list())

        tk.Button(rbar, text="📸 フォルダ", command=lambda: self.open_subfolder("screenshots"), font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=6, pady=3, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(rbar, text="↻", command=self.refresh_screenshot_accounts, font=("Yu Gothic UI", 8), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=6, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 4))

        self.ss_listbox = tk.Listbox(
            right, height=8, font=("Consolas", 9), bg="#0b1220", fg="#e2e8f0",
            selectbackground="#2563eb", selectforeground="#ffffff", borderwidth=0, highlightthickness=1, highlightbackground="#334155",
        )
        self.ss_listbox.pack(fill=tk.X, pady=(0, 6))
        self.ss_listbox.bind("<<ListboxSelect>>", lambda _e: self.preview_selected_screenshot())
        self.ss_listbox.bind("<Double-Button-1>", lambda _e: self.open_selected_screenshot())

        prev_frame = tk.Frame(right, bg="#020617", highlightthickness=1, highlightbackground="#334155")
        prev_frame.pack(fill=tk.BOTH, expand=True)
        self.ss_preview_lbl = tk.Label(
            prev_frame, text="一覧からスクショを選択すると\nここにプレビュー表示されます\n（ダブルクリックで実寸表示）",
            fg="#94a3b8", bg="#020617", font=("Yu Gothic UI", 9), justify=tk.CENTER
        )
        self.ss_preview_lbl.pack(fill=tk.BOTH, expand=True)

        self._ss_paths: list[str] = []

    def refresh_log_file_list(self):
        if not hasattr(self, "log_source_combo"):
            return
        items = ["(リアルタイム出力)"]
        if os.path.isdir(self.log_root):
            for fn in sorted(os.listdir(self.log_root)):
                if fn.endswith(".log"):
                    items.append(fn)
        self.log_source_combo["values"] = items
        if self.log_source_var.get() not in items:
            self.log_source_var.set("(リアルタイム出力)")

    def _insert_colored_log_line(self, line: str):
        tag = None
        if "[CRITICAL]" in line or "[ERROR]" in line or "エラー" in line or "凍結" in line:
            tag = "err"
        elif "[WARNING]" in line or "⚠️" in line or "リトライ" in line:
            tag = "warn"
        elif "送信完了" in line or "✅" in line or "posted" in line:
            tag = "ok"
        elif "📸" in line or "🧪" in line or "🔎" in line:
            tag = "info_hi"
        if tag:
            self.log_text.insert(tk.END, line, tag)
        else:
            self.log_text.insert(tk.END, line)

    def reload_log_view(self):
        src = self.log_source_var.get()
        if src == "(リアルタイム出力)":
            return
        path = os.path.join(self.log_root, src)
        self.log_text.delete("1.0", tk.END)
        if not os.path.isfile(path):
            return
        err_only = self.log_err_only_var.get()
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                lines = f.readlines()[-1000:]
            for ln in lines:
                if err_only and not any(k in ln for k in ("[ERROR]", "[CRITICAL]", "[WARNING]", "エラー", "凍結", "⚠️", "❌")):
                    continue
                self._insert_colored_log_line(ln)
            self.log_text.see(tk.END)
        except Exception as e:
            self.log_text.insert(tk.END, f"ログ読込エラー: {e}\n")

    def refresh_screenshot_accounts(self):
        if not hasattr(self, "ss_acc_combo"):
            return
        items = ["(すべて)"]
        if os.path.isdir(self.ss_root):
            for entry in sorted(os.listdir(self.ss_root)):
                if os.path.isdir(os.path.join(self.ss_root, entry)):
                    items.append(entry)
        self.ss_acc_combo["values"] = items
        if self.ss_acc_var.get() not in items:
            self.ss_acc_var.set("(すべて)")
        self.refresh_screenshot_list()

    def refresh_screenshot_list(self):
        self.ss_listbox.delete(0, tk.END)
        self._ss_paths = []
        acc_sel = self.ss_acc_var.get()
        if acc_sel == "(すべて)":
            pattern = os.path.join(self.ss_root, "*", "*.png")
        else:
            pattern = os.path.join(self.ss_root, acc_sel, "*.png")
        files = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)[:150]
        for fp in files:
            rel = os.path.relpath(fp, self.ss_root).replace("\\", "/")
            self._ss_paths.append(fp)
            self.ss_listbox.insert(tk.END, rel)

    def _load_image_into_label(self, path: str, label_widget: tk.Label, max_w: int = 380, max_h: int = 320, is_img_tab: bool = False):
        if not os.path.isfile(path):
            return
        try:
            photo = None
            if HAS_PIL:
                img = Image.open(path)
                img.thumbnail((max_w, max_h))
                photo = ImageTk.PhotoImage(img)
            elif path.lower().endswith((".png", ".gif")):
                raw_photo = tk.PhotoImage(file=path)
                w, h = raw_photo.width(), raw_photo.height()
                scale = max(1, (w + max_w - 1) // max_w, (h + max_h - 1) // max_h)
                photo = raw_photo.subsample(scale, scale) if scale > 1 else raw_photo

            if photo is not None:
                if is_img_tab:
                    self._img_preview_ref = photo
                else:
                    self._preview_photo_ref = photo
                label_widget.config(image=photo, text="")
            else:
                label_widget.config(image="", text=f"{os.path.basename(path)}\n(ダブルクリックで外部ビューア表示)")
        except Exception:
            label_widget.config(image="", text=f"{os.path.basename(path)}\n(プレビュー非対応・ダブルクリックで開く)")

    def preview_selected_screenshot(self):
        sel = self.ss_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self._ss_paths):
            self._load_image_into_label(self._ss_paths[idx], self.ss_preview_lbl, max_w=420, max_h=380, is_img_tab=False)

    def open_selected_screenshot(self):
        sel = self.ss_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self._ss_paths):
            self.open_path(self._ss_paths[idx])

    # --------------------------------------------------
    # Bot Start / Stop / Logging
    # --------------------------------------------------
    def append_log(self, line: str):
        self.log_queue.put(line)

    def _poll_log_queue(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                if self.log_source_var.get() == "(リアルタイム出力)":
                    if self.log_err_only_var.get():
                        if not any(k in line for k in ("[ERROR]", "[CRITICAL]", "[WARNING]", "エラー", "凍結", "⚠️", "❌")):
                            continue
                    self._insert_colored_log_line(line)
                    self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)

    def open_subfolder(self, folder_name: str):
        target = os.path.join(self.app_dir, folder_name)
        try:
            os.makedirs(target, exist_ok=True)
        except Exception:
            pass
        self.open_path(target)

    def _update_status_badge(self):
        internal_alive = self.bot_proc is not None and self.bot_proc.poll() is None
        if internal_alive:
            self.external_running = True
            self.status_badge.config(text="● ボット稼働中 (GUI監視)", fg="#bbf7d0", bg="#065f46")
        else:
            if self.bot_proc is not None:
                self.bot_proc = None
            ext = check_any_bot_running_fast(self.app_dir)
            self.external_running = ext
            if ext:
                self.status_badge.config(text="● ボット稼働中 (外部起動)", fg="#fde68a", bg="#92400e")
            else:
                self.status_badge.config(text="● 停止中", fg="#cbd5e1", bg="#334155")
        self._refresh_dashboard_countdowns_only()

    def _poll_bot_status(self):
        self._update_status_badge()
        self.root.after(1500, self._poll_bot_status)

    def _launch_bot_subprocess(self, cmd: list[str], start_banner: str, on_done=None):
        stop_flag = os.path.join(self.app_dir, "stop_bot.flag")
        if os.path.exists(stop_flag):
            try:
                os.remove(stop_flag)
            except Exception:
                pass
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        try:
            self.bot_proc = subprocess.Popen(
                cmd,
                cwd=self.app_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            self.log_source_var.set("(リアルタイム出力)")
            self.append_log(start_banner)
            threading.Thread(target=self._reader_thread, args=(self.bot_proc, on_done), daemon=True).start()
            self.notebook.select(self.tab_logs)
            self._update_status_badge()
        except Exception as e:
            messagebox.showerror("起動エラー", f"ボットの起動に失敗しました:\n{e}")

    def start_bot(self):
        if self.bot_proc is not None and self.bot_proc.poll() is None:
            messagebox.showinfo("稼働中", "すでにボットは稼働中です。")
            return
        if check_any_bot_running_fast(self.app_dir):
            if not messagebox.askyesno("確認", "すでに別の画面でボットが動いています。安全停止してからGUIから再起動しますか？"):
                return
            request_graceful_stop(self.app_dir)
            for _ in range(16):
                if not check_any_bot_running_fast(self.app_dir):
                    break
                time.sleep(0.25)
            if check_any_bot_running_fast(self.app_dir):
                kill_external_bots_on_windows()
                time.sleep(0.4)

        py_script = os.path.join(self.app_dir, "auto_tweet_playwright.py")
        exe_path = os.path.join(self.app_dir, "X_AutoTweetBot.exe")
        if os.path.isfile(py_script):
            cmd = [find_python_cli(), "-u", py_script]
        elif os.path.isfile(exe_path):
            cmd = [exe_path]
        else:
            messagebox.showerror("エラー", "auto_tweet_playwright.py も X_AutoTweetBot.exe も見つかりません。")
            return

        self._launch_bot_subprocess(cmd, "\n=== 自動投稿ボットを開始しました ===\n", on_done=self.refresh_all_views)

    def _reader_thread(self, proc: subprocess.Popen, on_done=None):
        try:
            if proc.stdout:
                for line in proc.stdout:
                    self.append_log(line)
        except Exception:
            pass
        rc = proc.wait()
        self.append_log(f"\n=== 処理終了 (終了コード: {rc}) ===\n")
        if on_done:
            self.root.after(150, on_done)

    def stop_bot(self):
        internal_alive = self.bot_proc is not None and self.bot_proc.poll() is None
        external_alive = check_any_bot_running_fast(self.app_dir)
        if not internal_alive and not external_alive:
            messagebox.showinfo("停止確認", "現在稼働中のボットはありません。（GUIや手動Chromeはそのまま開いた状態です）")
            return

        active_profile_before_stop = get_active_bot_profile(self.app_dir)
        self.btn_stop.config(state=tk.DISABLED, text="停止中...")
        self.status_badge.config(text="● 安全停止中...", fg="#fde68a", bg="#78350f")
        self.append_log("\n🛑 安全停止を指示しました（ボット側の処理区切りを待って静かに終了します）...\n")
        request_graceful_stop(self.app_dir)

        def _worker():
            observed_profile = active_profile_before_stop
            for _ in range(40):
                if not observed_profile:
                    observed_profile = get_active_bot_profile(self.app_dir)
                i_alive = self.bot_proc is not None and self.bot_proc.poll() is None
                e_alive = check_any_bot_running_fast(self.app_dir)
                if not i_alive and not e_alive:
                    break
                time.sleep(0.25)

            if self.bot_proc is not None and self.bot_proc.poll() is None:
                try:
                    self.bot_proc.terminate()
                    self.bot_proc.wait(timeout=3)
                except Exception:
                    try:
                        self.bot_proc.kill()
                    except Exception:
                        pass
            self.bot_proc = None

            if check_any_bot_running_fast(self.app_dir):
                kill_external_bots_on_windows()

            for fp in ("stop_bot.flag", "bot_running.pid", "bot_active_profile.txt"):
                p = os.path.join(self.app_dir, fp)
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass

            self.root.after(0, lambda prof=observed_profile: self._on_stop_finished(prof))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_stop_finished(self, stopped_bot_profile: str = ""):
        self.btn_stop.config(state=tk.NORMAL, text="■ ボット停止")
        self._update_status_badge()
        self.refresh_log_file_list()
        self.refresh_screenshot_accounts()
        self.refresh_dashboard()
        self.append_log("🛑 ボットを安全に停止しました。\n")
        if stopped_bot_profile:
            reopen = messagebox.askyesno(
                "停止完了",
                f"自動投稿ボットを安全に停止しました。\n\n"
                f"※ ボットが自動投稿用に開いていた「{stopped_bot_profile}」のブラウザ（画像非表示モード）は、プロファイルロック解除のため安全に終了しました。\n"
                f"（GUI画面や、手動で開いていた別プロファイルのChromeは閉じずに残ります）\n\n"
                f"このまま「{stopped_bot_profile}」を手動編集用の通常Chromeとして開き直しますか？",
            )
            if reopen:
                self._spawn_manual_chrome(stopped_bot_profile)
        else:
            messagebox.showinfo(
                "停止完了",
                "自動投稿ボットを安全に停止しました。\n\n※ このGUIウィンドウや、手動で開いたChromeウィンドウは閉じずにそのまま維持されます。",
            )

    # --------------------------------------------------
    # Chrome Profile Management Methods
    # --------------------------------------------------
    def refresh_profiles(self):
        if not hasattr(self, "profile_tree"):
            return
        for iid in self.profile_tree.get_children():
            self.profile_tree.delete(iid)
        if not os.path.isdir(self.sessions_root):
            return
        sess_status = read_json_safe(os.path.join(self.app_dir, "session_status.json"), {})
        if not isinstance(sess_status, dict):
            sess_status = {}
        for entry in sorted(os.listdir(self.sessions_root)):
            if os.path.isdir(os.path.join(self.sessions_root, entry)):
                cid = entry.lstrip("_").lower()
                st_obj = sess_status.get(cid, {})
                st_txt = st_obj.get("status", "未チェック") if isinstance(st_obj, dict) else "未チェック"
                chk_txt = st_obj.get("checked_at", "-") if isinstance(st_obj, dict) else "-"
                self.profile_tree.insert("", tk.END, values=(entry, st_txt, chk_txt))

    def on_select_profile(self, _event=None):
        sel = self.profile_tree.selection()
        if not sel:
            return
        vals = self.profile_tree.item(sel[0], "values")
        if vals:
            self.profile_name_var.set(str(vals[0]).lstrip("_"))

    def _resolve_profile_dir(self, clean_name: str) -> tuple[str, bool]:
        exact_dir = os.path.join(self.sessions_root, clean_name)
        underscore_dir = os.path.join(self.sessions_root, f"_{clean_name}")
        if os.path.exists(exact_dir):
            return os.path.abspath(exact_dir), True
        if os.path.exists(underscore_dir):
            return os.path.abspath(underscore_dir), True
        os.makedirs(exact_dir, exist_ok=True)
        return os.path.abspath(exact_dir), False

    def _spawn_manual_chrome(self, clean_name: str):
        chrome_exe = find_chrome_executable()
        if not chrome_exe:
            messagebox.showerror("エラー", "Google Chrome が見つかりませんでした。")
            return

        profile_dir, existed = self._resolve_profile_dir(clean_name)
        self.refresh_profiles()
        creationflags = 0
        if sys.platform == "win32":
            creationflags = (
                getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
                | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
            )
        try:
            subprocess.Popen(
                [chrome_exe, f"--user-data-dir={profile_dir}"],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
                creationflags=creationflags,
            )
            msg = "既存プロファイル" if existed else "新規プロファイル"
            self.append_log(f"🌐 [{clean_name}] ({msg}) を独立した手動編集用Chromeで開きました。\n")
        except Exception as e:
            messagebox.showerror("Chrome起動エラー", f"Chromeを起動できませんでした:\n{e}")

    def launch_chrome_for_profile(self):
        raw = self.profile_name_var.get().strip()
        if not raw:
            sel = self.profile_tree.selection()
            if sel:
                raw = str(self.profile_tree.item(sel[0], "values")[0])
        clean = sanitize_profile_name(raw)
        if not clean:
            messagebox.showwarning("入力確認", "プロファイル名を入力するか、一覧から選択してください。")
            return

        bot_alive = (self.bot_proc is not None and self.bot_proc.poll() is None) or check_any_bot_running_fast(self.app_dir)
        active_prof = get_active_bot_profile(self.app_dir) if bot_alive else ""

        if bot_alive:
            if active_prof and active_prof.lower() == clean.lower():
                if not messagebox.askyesno(
                    "同一プロファイル使用中",
                    f"現在ボットが「{clean}」のChromeを使用中です。\n"
                    f"ボットを止めずに同じプロファイルを開くとChromeが統合され、ボット停止時に一緒に閉じてしまいます。\n\n"
                    f"ボットを安全停止してから、独立した手動Chromeで開きますか？",
                ):
                    return
                stop_now = True
            else:
                ans = messagebox.askyesnocancel(
                    "ボット稼働中の確認",
                    "現在自動投稿ボットが稼働中です。\nボットを安全停止してからChromeを開きますか？\n\n"
                    "「はい」= ボットを安全停止してChromeを開く\n"
                    "「いいえ」= ボットを止めずに別プロファイルとしてChromeを開く",
                )
                if ans is None:
                    return
                stop_now = bool(ans)

            if stop_now:
                request_graceful_stop(self.app_dir)
                for _ in range(32):
                    i_alive = self.bot_proc is not None and self.bot_proc.poll() is None
                    if not i_alive and not check_any_bot_running_fast(self.app_dir):
                        break
                    time.sleep(0.25)
                if self.bot_proc is not None and self.bot_proc.poll() is None:
                    self.bot_proc.terminate()
                    self.bot_proc = None
                if check_any_bot_running_fast(self.app_dir):
                    kill_external_bots_on_windows()
                for fp in ("stop_bot.flag", "bot_running.pid", "bot_active_profile.txt"):
                    p = os.path.join(self.app_dir, fp)
                    try:
                        if os.path.exists(p):
                            os.remove(p)
                    except Exception:
                        pass
                self.append_log("\n🛑 プロファイル編集のためボットを安全停止しました。\n")
                self._update_status_badge()

        self._spawn_manual_chrome(clean)

    def rename_selected_profile(self):
        sel = self.profile_tree.selection()
        if not sel:
            messagebox.showinfo("選択", "名前変更するプロファイルを一覧から選択してください。")
            return
        old_name = str(self.profile_tree.item(sel[0], "values")[0])
        new_raw = simpledialog.askstring("プロファイル名変更", f"「{old_name}」の新しいプロファイル名:", initialvalue=old_name, parent=self.root)
        if not new_raw:
            return
        new_name = sanitize_profile_name(new_raw)
        if not new_name or new_name == old_name:
            return
        src = os.path.join(self.sessions_root, old_name)
        dst = os.path.join(self.sessions_root, new_name)
        if os.path.exists(dst):
            messagebox.showerror("重複", f"「{new_name}」はすでに存在します。")
            return
        try:
            os.rename(src, dst)
            self.profile_name_var.set(new_name)
            self.refresh_profiles()
        except Exception as e:
            messagebox.showerror("変更エラー", f"開いているChromeを閉じてから再実行してください。\n{e}")

    def delete_selected_profile(self):
        sel = self.profile_tree.selection()
        if not sel:
            messagebox.showinfo("選択", "削除するプロファイルを一覧から選択してください。")
            return
        name = str(self.profile_tree.item(sel[0], "values")[0])
        if not messagebox.askyesno("削除確認", f"プロファイルフォルダ「{name}」を削除しますか？\n（このプロファイルのログイン情報が消去されます）"):
            return
        target = os.path.join(self.sessions_root, name)
        try:
            shutil.rmtree(target)
            if self.profile_name_var.get().strip() == name.lstrip("_"):
                self.profile_name_var.set("")
            self.refresh_profiles()
        except Exception as e:
            messagebox.showerror("削除エラー", f"Chromeを閉じてから再実行してください。\n{e}")

    # --------------------------------------------------
    # Accounts & Excel Management Methods
    # --------------------------------------------------
    def _find_account_excel_name(self, uid: str) -> str:
        clean = uid.replace("@", "").strip()
        for c in (f"{clean}.xlsx", f"@{clean}.xlsx", f"tweets_{clean}.xlsx"):
            if os.path.exists(os.path.join(self.app_dir, c)):
                return c
        return "(未作成)"

    def refresh_accounts(self):
        if not hasattr(self, "acc_tree"):
            return
        for item in self.acc_tree.get_children():
            self.acc_tree.delete(item)
        if not HAS_OPENPYXL:
            return
        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        if not os.path.exists(acc_path):
            return
        paused = load_paused_set(self.app_dir)
        suspended = {e["clean_id"] for e in load_suspended_entries(self.app_dir)}
        try:
            wb = openpyxl.load_workbook(acc_path, data_only=True)
            ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
            for r in range(2, ws.max_row + 1):
                uid = str(ws.cell(r, 1).value or "").strip()
                pw = str(ws.cell(r, 2).value or "").strip()
                nm = str(ws.cell(r, 3).value or "").strip()
                tok = str(ws.cell(r, 4).value or "").strip()
                if not uid and not nm:
                    continue
                clean_id = sanitize_profile_name(uid).lower()
                state = "🚫除外" if clean_id in suspended else ("⏸休止" if clean_id in paused else "🟢稼働")
                masked_pw = "●" * min(len(pw), 8) if pw else ""
                ex_name = self._find_account_excel_name(uid)
                self.acc_tree.insert("", tk.END, values=(uid, masked_pw, nm, tok[:18] + ("..." if len(tok) > 18 else ""), state, ex_name), tags=(str(r),))
        except Exception as e:
            self.append_log(f"⚠️ アカウント情報.xlsx 読込エラー: {e}\n")

    def on_select_account(self, _event=None):
        sel = self.acc_tree.selection()
        if not sel or not HAS_OPENPYXL:
            return
        vals = self.acc_tree.item(sel[0], "values")
        if not vals:
            return
        uid = vals[0]
        self.acc_id_var.set(uid)
        self.profile_name_var.set(uid.lstrip("@"))
        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        try:
            wb = openpyxl.load_workbook(acc_path, data_only=True)
            ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, 1).value or "").strip() == uid:
                    self.acc_pw_var.set(str(ws.cell(r, 2).value or "").strip())
                    self.acc_name_var.set(str(ws.cell(r, 3).value or "").strip())
                    self.acc_tok_var.set(str(ws.cell(r, 4).value or "").strip())
                    break
        except Exception:
            pass

    def save_account_row(self):
        if not HAS_OPENPYXL:
            messagebox.showerror("エラー", "openpyxl が読み込めないため編集できません。")
            return
        uid = self.acc_id_var.get().strip()
        if not uid:
            messagebox.showwarning("入力確認", "アカウントIDを入力してください。")
            return
        pw = self.acc_pw_var.get().strip()
        nm = self.acc_name_var.get().strip() or uid
        tok = self.acc_tok_var.get().strip()

        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        try:
            if os.path.exists(acc_path):
                wb = openpyxl.load_workbook(acc_path)
                ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "アカウント情報"
                for idx, h in enumerate(["アカウントID", "アカウントpw", "アカウント名", "auth_token"], 1):
                    ws.cell(1, idx, h)

            target_row = None
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, 1).value or "").strip() == uid:
                    target_row = r
                    break
            if target_row is None:
                target_row = max(2, ws.max_row + 1)
                while target_row > 2 and not any(ws.cell(target_row - 1, c).value for c in range(1, 5)):
                    target_row -= 1

            ws.cell(target_row, 1, uid)
            ws.cell(target_row, 2, pw)
            ws.cell(target_row, 3, nm)
            ws.cell(target_row, 4, tok)
            wb.save(acc_path)
            self.refresh_accounts()
            self.refresh_dashboard()
            messagebox.showinfo("保存完了", f"アカウント「{uid}」を アカウント情報.xlsx に保存しました。")
        except PermissionError:
            messagebox.showerror("保存エラー", "アカウント情報.xlsx が Excel で開かれているため保存できません。Excel を閉じて再実行してください。")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))

    def delete_account_row(self):
        if not HAS_OPENPYXL:
            return
        uid = self.acc_id_var.get().strip()
        if not uid:
            messagebox.showinfo("選択", "削除するアカウントを選択してください。")
            return
        if not messagebox.askyesno("削除確認", f"アカウント情報.xlsx から「{uid}」を削除しますか？"):
            return
        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        try:
            wb = openpyxl.load_workbook(acc_path)
            ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, 1).value or "").strip() == uid:
                    ws.delete_rows(r)
                    wb.save(acc_path)
                    break
            self.refresh_accounts()
            self.refresh_dashboard()
        except PermissionError:
            messagebox.showerror("削除エラー", "アカウント情報.xlsx を閉じてから実行してください。")
        except Exception as e:
            messagebox.showerror("削除エラー", str(e))

    def refresh_excel_files(self):
        if not hasattr(self, "excel_listbox"):
            return
        self.excel_listbox.delete(0, tk.END)
        for path in sorted(glob.glob(os.path.join(self.app_dir, "*.xlsx"))):
            fname = os.path.basename(path)
            if fname.startswith("~$"):
                continue
            self.excel_listbox.insert(tk.END, fname)

    def open_selected_excel(self):
        sel = self.excel_listbox.curselection()
        if not sel:
            messagebox.showinfo("選択", "開くExcelファイルを一覧から選択してください。")
            return
        fname = self.excel_listbox.get(sel[0])
        self.open_path(os.path.join(self.app_dir, fname))

    def run_excel_theme_updater(self):
        updater = os.path.join(self.app_dir, "update_themes_excel.py")
        if not os.path.isfile(updater):
            messagebox.showerror("エラー", "update_themes_excel.py が見つかりません。")
            return
        try:
            res = subprocess.run(
                [find_python_cli(), updater],
                cwd=self.app_dir,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            self.refresh_excel_files()
            self.refresh_theme_account_list()
            self.refresh_dashboard()
            out = (res.stdout or "") + (res.stderr or "")
            messagebox.showinfo("Excel最適化完了", out.strip() or "全Excelファイルのテーマ別カラー書式を更新しました。")
        except Exception as e:
            messagebox.showerror("実行エラー", str(e))

    def open_path(self, path: str):
        if not os.path.exists(path):
            messagebox.showwarning("見つかりません", f"対象が存在しません:\n{path}")
            return
        try:
            if sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def on_close(self):
        if self.bot_proc is not None and self.bot_proc.poll() is None:
            ans = messagebox.askyesnocancel(
                "終了確認",
                "自動投稿ボットが稼働中です。\nボットも停止してコントロールパネルを閉じますか？\n\n「はい」= ボットを停止して終了\n「いいえ」= ボットを動かしたままGUIだけ閉じる",
            )
            if ans is None:
                return
            if ans is True:
                request_graceful_stop(self.app_dir)
                try:
                    self.bot_proc.wait(timeout=3)
                except Exception:
                    try:
                        self.bot_proc.terminate()
                    except Exception:
                        pass
                kill_external_bots_on_windows()
        self.root.destroy()


def main():
    root = tk.Tk()
    XBotControlPanel(root)
    root.mainloop()


if __name__ == "__main__":
    main()
