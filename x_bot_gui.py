#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
X Auto Tweet Bot コントロールパネル (GUI)
- ボットの起動・リアルタイムログ表示・ワンクリック停止（外部起動プロセスも停止可能）
- Chromeプロファイル管理（作成 / 一覧 / リネーム / 削除 / 手動Chrome起動でXプロフィール編集・再ログイン）
- アカウント情報.xlsx のGUI編集 & テーマ別Excelファイル管理
"""

import os
import re
import sys
import glob
import time
import queue
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog

try:
    import openpyxl
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


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
    """GUIからサブプロセスを実行するための console 版 python.exe を解決"""
    exe = sys.executable
    if exe and os.path.basename(exe).lower() == "pythonw.exe":
        py_cli = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.isfile(py_cli):
            return py_cli
    return exe or "python"


def sanitize_profile_name(raw_name: str) -> str:
    cleaned = raw_name.strip().lstrip("@").strip()
    return re.sub(r'[<>:"/\\|?*]', "_", cleaned)


def kill_external_bots_on_windows() -> int:
    """外部起動されたボットプロセスも安全に検出・停止"""
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
    """PowerShell起動なしで高速に外部ボット稼働有無を判定（ウィンドウちらつき・フォーカス移動なし）"""
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


class XBotControlPanel:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.app_dir = get_app_dir()
        os.chdir(self.app_dir)
        self.sessions_root = os.path.join(self.app_dir, "browser_sessions")
        os.makedirs(self.sessions_root, exist_ok=True)

        self.bot_proc: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.external_running = False

        self.root.title("X Auto Tweet Bot コントロールパネル (GUI)")
        self.root.geometry("980x700")
        self.root.minsize(820, 580)
        self.root.configure(bg="#0f172a")

        self._setup_styles()
        self._build_ui()

        self.refresh_profiles()
        self.refresh_accounts()
        self.refresh_excel_files()

        self.root.after(150, self._poll_log_queue)
        self.root.after(1500, self._poll_bot_status)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        default_font = ("Yu Gothic UI", 10)
        style.configure(".", font=default_font)
        style.configure("TNotebook", background="#0f172a", borderwidth=0)
        style.configure("TNotebook.Tab", background="#1e293b", foreground="#cbd5e1", padding=[14, 7], font=("Yu Gothic UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", "#1e3a8a")], foreground=[("selected", "#ffffff")])
        style.configure("Card.TFrame", background="#1e293b")
        style.configure("Dark.TFrame", background="#0f172a")
        style.configure("Treeview", background="#0b1220", foreground="#e2e8f0", fieldbackground="#0b1220", rowheight=26, font=("Yu Gothic UI", 9))
        style.configure("Treeview.Heading", background="#1e3a8a", foreground="#ffffff", font=("Yu Gothic UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#2563eb")])

    def _build_ui(self):
        # ===== Top Control Bar =====
        header = tk.Frame(self.root, bg="#111827", pady=12, padx=16, highlightthickness=1, highlightbackground="#1f2937")
        header.pack(fill=tk.X)

        left_hdr = tk.Frame(header, bg="#111827")
        left_hdr.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(
            left_hdr,
            text="X Auto Tweet Bot コントロールパネル",
            font=("Yu Gothic UI", 14, "bold"),
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

        tk.Label(
            info_row,
            text=f"  作業フォルダ: {self.app_dir}",
            font=("Yu Gothic UI", 9),
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
            pady=7,
            cursor="hand2",
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

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
            pady=7,
            cursor="hand2",
        )
        self.btn_stop.pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(
            btn_bar,
            text="📁 フォルダを開く",
            command=lambda: self.open_path(self.app_dir),
            font=("Yu Gothic UI", 9, "bold"),
            bg="#334155",
            fg="#e2e8f0",
            activebackground="#475569",
            relief=tk.FLAT,
            padx=10,
            pady=7,
            cursor="hand2",
        ).pack(side=tk.LEFT)

        # ===== Main Tabs =====
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.tab_profiles = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_accounts = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.tab_logs = ttk.Frame(self.notebook, style="Dark.TFrame")

        self.notebook.add(self.tab_profiles, text="  👤 Chromeプロファイル管理・編集  ")
        self.notebook.add(self.tab_accounts, text="  📊 アカウント情報・Excel管理  ")
        self.notebook.add(self.tab_logs, text="  🖥️ ボット稼働ログ  ")

        self._build_profiles_tab()
        self._build_accounts_tab()
        self._build_logs_tab()

    # --------------------------------------------------
    # Tab 1: Chromeプロファイル管理・編集
    # --------------------------------------------------
    def _build_profiles_tab(self):
        container = tk.Frame(self.tab_profiles, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        # Left column: Existing profiles list
        left = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        hdr = tk.Frame(left, bg="#1e293b")
        hdr.pack(fill=tk.X, pady=(0, 8))
        tk.Label(hdr, text="保存済みプロファイル一覧 (browser_sessions)", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(hdr, text="↻ 更新", command=self.refresh_profiles, font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=2, cursor="hand2").pack(side=tk.RIGHT)

        list_frame = tk.Frame(left, bg="#1e293b")
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.profile_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Consolas", 11),
            bg="#0b1220",
            fg="#f8fafc",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#334155",
            activestyle="none",
        )
        self.profile_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.profile_listbox.yview)
        self.profile_listbox.bind("<<ListboxSelect>>", self.on_select_profile)
        self.profile_listbox.bind("<Double-Button-1>", lambda _e: self.launch_chrome_for_profile())

        btn_row_left = tk.Frame(left, bg="#1e293b")
        btn_row_left.pack(fill=tk.X, pady=(10, 0))

        tk.Button(
            btn_row_left,
            text="✏️ 名前変更",
            command=self.rename_selected_profile,
            font=("Yu Gothic UI", 9, "bold"),
            bg="#334155",
            fg="#f8fafc",
            relief=tk.FLAT,
            padx=10,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btn_row_left,
            text="🗑️ 選択削除",
            command=self.delete_selected_profile,
            font=("Yu Gothic UI", 9, "bold"),
            bg="#7f1d1d",
            fg="#fecaca",
            relief=tk.FLAT,
            padx=10,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.LEFT, padx=(0, 6))

        tk.Button(
            btn_row_left,
            text="📂 保存先フォルダを開く",
            command=lambda: self.open_path(self.sessions_root),
            font=("Yu Gothic UI", 9),
            bg="#334155",
            fg="#cbd5e1",
            relief=tk.FLAT,
            padx=10,
            pady=6,
            cursor="hand2",
        ).pack(side=tk.RIGHT)

        # Right column: Launch & edit panel
        right = tk.Frame(container, bg="#1e293b", padx=16, pady=16, width=380, highlightthickness=1, highlightbackground="#334155")
        right.pack(side=tk.RIGHT, fill=tk.Y)
        right.pack_propagate(False)

        tk.Label(right, text="Chrome 手動起動（ログイン・Xプロフィール編集）", font=("Yu Gothic UI", 11, "bold"), fg="#38bdf8", bg="#1e293b", wraplength=340, justify="left").pack(anchor="w")

        tk.Label(
            right,
            text="・一覧からプロファイルを選ぶか、新しい名前を入力してChromeを起動します。\n"
                 "・サイトは自動では開きません。ご自身でXを開いてプロフィール編集・アイコン変更・ログイン等を行えます。\n"
                 "・開いたChromeを「×」で閉じるとCookie/セッションが保存され、自動投稿ボットに引き継がれます。",
            font=("Yu Gothic UI", 9),
            fg="#cbd5e1",
            bg="#1e293b",
            justify="left",
            wraplength=340,
        ).pack(anchor="w", pady=(8, 14))

        tk.Label(right, text="プロファイル名（@なしID推奨 例: wqwq1320）", font=("Yu Gothic UI", 9, "bold"), fg="#e2e8f0", bg="#1e293b").pack(anchor="w")
        self.profile_name_var = tk.StringVar()
        self.entry_profile_name = tk.Entry(
            right,
            textvariable=self.profile_name_var,
            font=("Consolas", 11),
            bg="#0b1220",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#475569",
        )
        self.entry_profile_name.pack(fill=tk.X, ipady=6, pady=(4, 12))

        tk.Button(
            right,
            text="🌐 このプロファイルで Chrome を開く",
            command=self.launch_chrome_for_profile,
            font=("Yu Gothic UI", 10, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0ea5e9",
            relief=tk.FLAT,
            pady=10,
            cursor="hand2",
        ).pack(fill=tk.X, pady=(0, 10))

        self.chrome_status_lbl = tk.Label(
            right,
            text="※ ボット稼働中に同じプロファイルを開くと競合するため、自動で一時停止確認を行います。",
            font=("Yu Gothic UI", 8),
            fg="#fbbf24",
            bg="#1e293b",
            wraplength=340,
            justify="left",
        )
        self.chrome_status_lbl.pack(anchor="w", pady=(6, 0))

    # --------------------------------------------------
    # Tab 2: アカウント情報・Excel管理
    # --------------------------------------------------
    def _build_accounts_tab(self):
        container = tk.Frame(self.tab_accounts, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        # Top: アカウント情報.xlsx エディタ
        top_card = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        top_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        hdr = tk.Frame(top_card, bg="#1e293b")
        hdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(hdr, text="アカウント情報.xlsx 管理", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(hdr, text="📊 アカウント情報.xlsx をExcelで開く", command=lambda: self.open_path(os.path.join(self.app_dir, "アカウント情報.xlsx")), font=("Yu Gothic UI", 9), bg="#065f46", fg="#d1fae5", relief=tk.FLAT, padx=10, pady=3, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(hdr, text="↻ 再読込", command=self.refresh_accounts, font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        cols = ("id", "pw", "name", "token", "excel")
        self.acc_tree = ttk.Treeview(top_card, columns=cols, show="headings", height=6)
        self.acc_tree.heading("id", text="アカウントID")
        self.acc_tree.heading("pw", text="アカウントpw")
        self.acc_tree.heading("name", text="アカウント名")
        self.acc_tree.heading("token", text="auth_token")
        self.acc_tree.heading("excel", text="対応Excel")
        self.acc_tree.column("id", width=160)
        self.acc_tree.column("pw", width=130)
        self.acc_tree.column("name", width=150)
        self.acc_tree.column("token", width=190)
        self.acc_tree.column("excel", width=160)
        self.acc_tree.pack(fill=tk.BOTH, expand=True)
        self.acc_tree.bind("<<TreeviewSelect>>", self.on_select_account)

        # Inline add / update form
        form = tk.Frame(top_card, bg="#1e293b", pady=8)
        form.pack(fill=tk.X)

        self.acc_id_var = tk.StringVar()
        self.acc_pw_var = tk.StringVar()
        self.acc_name_var = tk.StringVar()
        self.acc_tok_var = tk.StringVar()

        fields = [
            ("ID:", self.acc_id_var, 16),
            ("PW:", self.acc_pw_var, 14),
            ("表示名:", self.acc_name_var, 16),
            ("auth_token:", self.acc_tok_var, 20),
        ]
        for lbl, var, w in fields:
            tk.Label(form, text=lbl, fg="#cbd5e1", bg="#1e293b", font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, padx=(0, 3))
            tk.Entry(form, textvariable=var, width=w, bg="#0b1220", fg="#ffffff", insertbackground="#ffffff", relief=tk.FLAT, highlightthickness=1, highlightbackground="#475569").pack(side=tk.LEFT, padx=(0, 8), ipady=3)

        tk.Button(form, text="＋ 追加 / 上書き保存", command=self.save_account_row, font=("Yu Gothic UI", 9, "bold"), bg="#2563eb", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(form, text="削除", command=self.delete_account_row, font=("Yu Gothic UI", 9, "bold"), bg="#7f1d1d", fg="#fecaca", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.LEFT)

        # Bottom: Excelファイル一覧 & テーマ最適化
        bot_card = tk.Frame(container, bg="#1e293b", padx=12, pady=12, highlightthickness=1, highlightbackground="#334155")
        bot_card.pack(fill=tk.BOTH, expand=True)

        bhdr = tk.Frame(bot_card, bg="#1e293b")
        bhdr.pack(fill=tk.X, pady=(0, 6))
        tk.Label(bhdr, text="投稿用 Excelファイル一覧 (テーマ別シート)", font=("Yu Gothic UI", 10, "bold"), fg="#ffffff", bg="#1e293b").pack(side=tk.LEFT)
        tk.Button(bhdr, text="🎨 全Excelをテーマ形式＆色付けに更新", command=self.run_excel_theme_updater, font=("Yu Gothic UI", 9, "bold"), bg="#d97706", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(bhdr, text="📊 選択したExcelを開く", command=self.open_selected_excel, font=("Yu Gothic UI", 9, "bold"), bg="#059669", fg="#ffffff", relief=tk.FLAT, padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        self.excel_listbox = tk.Listbox(
            bot_card,
            height=5,
            font=("Consolas", 10),
            bg="#0b1220",
            fg="#e2e8f0",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#334155",
        )
        self.excel_listbox.pack(fill=tk.BOTH, expand=True)
        self.excel_listbox.bind("<Double-Button-1>", lambda _e: self.open_selected_excel())

    # --------------------------------------------------
    # Tab 3: ボット稼働ログ
    # --------------------------------------------------
    def _build_logs_tab(self):
        container = tk.Frame(self.tab_logs, bg="#0f172a", padx=8, pady=8)
        container.pack(fill=tk.BOTH, expand=True)

        bar = tk.Frame(container, bg="#0f172a")
        bar.pack(fill=tk.X, pady=(0, 6))
        tk.Label(bar, text="リアルタイム実行ログ (log/<アカウント名>.log ＆ screenshots/<アカウント名>/)", font=("Yu Gothic UI", 10, "bold"), fg="#cbd5e1", bg="#0f172a").pack(side=tk.LEFT)
        tk.Button(bar, text="📂 logフォルダを開く", command=lambda: self.open_subfolder("log"), font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(bar, text="📸 スクショフォルダを開く", command=lambda: self.open_subfolder("screenshots"), font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))
        tk.Button(bar, text="ログ消去", command=lambda: self.log_text.delete("1.0", tk.END), font=("Yu Gothic UI", 9), bg="#334155", fg="#e2e8f0", relief=tk.FLAT, padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=(0, 6))

        self.log_text = scrolledtext.ScrolledText(
            container,
            bg="#020617",
            fg="#e2e8f0",
            font=("Consolas", 10),
            insertbackground="#ffffff",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#334155",
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # --------------------------------------------------
    # Bot Start / Stop / Logging
    # --------------------------------------------------
    def append_log(self, line: str):
        self.log_queue.put(line)

    def _poll_log_queue(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, line)
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

    def _poll_bot_status(self):
        self._update_status_badge()
        self.root.after(1500, self._poll_bot_status)

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

        stop_flag = os.path.join(self.app_dir, "stop_bot.flag")
        if os.path.exists(stop_flag):
            try:
                os.remove(stop_flag)
            except Exception:
                pass

        py_script = os.path.join(self.app_dir, "auto_tweet_playwright.py")
        exe_path = os.path.join(self.app_dir, "X_AutoTweetBot.exe")
        if os.path.isfile(py_script):
            cmd = [find_python_cli(), "-u", py_script]
        elif os.path.isfile(exe_path):
            cmd = [exe_path]
        else:
            messagebox.showerror("エラー", "auto_tweet_playwright.py も X_AutoTweetBot.exe も見つかりません。")
            return

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
            self.append_log("\n=== 自動投稿ボットを開始しました ===\n")
            threading.Thread(target=self._reader_thread, args=(self.bot_proc,), daemon=True).start()
            self.notebook.select(self.tab_logs)
            self._update_status_badge()
        except Exception as e:
            messagebox.showerror("起動エラー", f"ボットの起動に失敗しました:\n{e}")

    def _reader_thread(self, proc: subprocess.Popen):
        try:
            if proc.stdout:
                for line in proc.stdout:
                    self.append_log(line)
        except Exception:
            pass
        rc = proc.wait()
        self.append_log(f"\n=== ボット停止 (終了コード: {rc}) ===\n")

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
            # stop_bot.flag によるクリーン終了を最大10秒待機（Excel保存やcontext.close()を安全完了させる）
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
    # Chrome Profile Management
    # --------------------------------------------------
    def refresh_profiles(self):
        self.profile_listbox.delete(0, tk.END)
        if not os.path.isdir(self.sessions_root):
            return
        for entry in sorted(os.listdir(self.sessions_root)):
            if os.path.isdir(os.path.join(self.sessions_root, entry)):
                self.profile_listbox.insert(tk.END, entry)

    def on_select_profile(self, _event=None):
        sel = self.profile_listbox.curselection()
        if not sel:
            return
        name = self.profile_listbox.get(sel[0])
        self.profile_name_var.set(name.lstrip("_"))

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
            sel = self.profile_listbox.curselection()
            if sel:
                raw = self.profile_listbox.get(sel[0])
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
        sel = self.profile_listbox.curselection()
        if not sel:
            messagebox.showinfo("選択", "名前変更するプロファイルを一覧から選択してください。")
            return
        old_name = self.profile_listbox.get(sel[0])
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
        sel = self.profile_listbox.curselection()
        if not sel:
            messagebox.showinfo("選択", "削除するプロファイルを一覧から選択してください。")
            return
        name = self.profile_listbox.get(sel[0])
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
    # Accounts & Excel Management
    # --------------------------------------------------
    def _find_account_excel_name(self, uid: str) -> str:
        clean = uid.replace("@", "").strip()
        for c in (f"{clean}.xlsx", f"@{clean}.xlsx", f"tweets_{clean}.xlsx"):
            if os.path.exists(os.path.join(self.app_dir, c)):
                return c
        return "(未作成)"

    def refresh_accounts(self):
        for item in self.acc_tree.get_children():
            self.acc_tree.delete(item)
        if not HAS_OPENPYXL:
            return
        acc_path = os.path.join(self.app_dir, "アカウント情報.xlsx")
        if not os.path.exists(acc_path):
            return
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
                masked_pw = "●" * min(len(pw), 8) if pw else ""
                ex_name = self._find_account_excel_name(uid)
                self.acc_tree.insert("", tk.END, values=(uid, masked_pw, nm, tok[:18] + ("..." if len(tok) > 18 else ""), ex_name), tags=(str(r),))
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
        except PermissionError:
            messagebox.showerror("削除エラー", "アカウント情報.xlsx を閉じてから実行してください。")
        except Exception as e:
            messagebox.showerror("削除エラー", str(e))

    def refresh_excel_files(self):
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
                try:
                    self.bot_proc.terminate()
                except Exception:
                    pass
                kill_external_bots_on_windows()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        XBotControlPanel(root)
        root.mainloop()
    except Exception as e:
        root.withdraw()
        messagebox.showerror("GUI 起動エラー", f"エラーが発生しました:\n{e}")


if __name__ == "__main__":
    main()
