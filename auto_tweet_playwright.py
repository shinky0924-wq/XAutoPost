#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
X (Twitter) ブラウザ自動操縦・自動ツイートボット (テーマ別管理・完全統合版)
====================================================================
【統合機能一覧】
  1. 【テーマ別マルチシート管理】
     - Excelの下部シート（タブ）ごとに独立した「テーマ」として運用
     - 各シート1〜2行目に「実行時間 / 曜日 / ポスト間隔 / 投稿モード / 有効/無効」を設定
  2. 【ポスト別 投稿時間フィルタ＆即スキップ】
     - 各ポストの「投稿時間」列（7～10, ７～１０, 7時から10時, 22～2, 7～10, 18～21 等）を判定
     - 時間外のポストはステータスを変更せず即スキップして次のポストへ進む
     - 空白のポストは通常通りテーマの実行時間・曜日に従って稼働
  3. 【1窓・順番交代ローテーション & メディア通信ブロック】
     - 常に開くブラウザは1つのみ ＆ アップロード通信を保護しつつ画像・動画読込を遮断
  4. 【スピンタックス＆動的タグ置換】
     - {A|B|C} ＆ {日付}{月日}{曜日}{時間}{時刻}{午前午後}
  5. 【案3: 画像クールダウン＆アカウント間重複防止】
     - 同一アカウント：30日間再選出禁止 (image_history.json)
     - 他アカウント：7日間共有禁止（同一IP重複対策）
  6. 【凍結アカウント即時永久除外】
     - 凍結・制限検知時にメモリ＆アカウント情報Excelから即座に除外し再起動不要で継続
  7. 【chrome_launcher.exe 連携 & 括弧付きフォルダ対応】
     - browser_sessions/ の手動ログイン済みプロファイルを自動引き継ぎ
     - X_BOT_配布用(Python不要) 等でも Playwright ドライバを正常起動
====================================================================
"""

import os
import sys
import re
import time
import glob
import json
import random
import logging
import asyncio
import tempfile
import traceback
import subprocess
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill
from datetime import datetime, timedelta, time as dt_time
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def sanitize_filename_part(name: str) -> str:
    cleaned = str(name or "").replace("@", "").strip()
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", cleaned)
    return cleaned or "unknown"


def get_log_dir() -> str:
    d = os.path.join(get_app_dir(), "log")
    os.makedirs(d, exist_ok=True)
    return d


def get_screenshot_dir(account_name: str | None = None) -> str:
    base = os.path.join(get_app_dir(), "screenshots")
    if account_name:
        target = os.path.join(base, sanitize_filename_part(account_name))
    else:
        target = base
    os.makedirs(target, exist_ok=True)
    return target


_LOG_FORMATTER = logging.Formatter('[%(asctime)s] [%(levelname)s] [XBrowserBot] %(message)s')
_CURRENT_FILE_HANDLER: logging.FileHandler | None = None
_CURRENT_LOG_PATH: str | None = None

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] [XBrowserBot] %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("XBrowserBot")


def switch_account_log(account_name: str | None = None):
    """出力先ファイルを log/<アカウント名>.log（待機・システム時は log/system.log）へ切り替える"""
    global _CURRENT_FILE_HANDLER, _CURRENT_LOG_PATH
    try:
        filename = f"{sanitize_filename_part(account_name)}.log" if account_name else "system.log"
        target_path = os.path.abspath(os.path.join(get_log_dir(), filename))
        if _CURRENT_LOG_PATH == target_path and _CURRENT_FILE_HANDLER is not None:
            return

        root_logger = logging.getLogger()
        if _CURRENT_FILE_HANDLER is not None:
            try:
                _CURRENT_FILE_HANDLER.flush()
                _CURRENT_FILE_HANDLER.close()
            except Exception:
                pass
            try:
                root_logger.removeHandler(_CURRENT_FILE_HANDLER)
            except Exception:
                pass
            _CURRENT_FILE_HANDLER = None

        fh = logging.FileHandler(target_path, encoding="utf-8")
        fh.setLevel(logging.INFO)
        fh.setFormatter(_LOG_FORMATTER)
        root_logger.addHandler(fh)
        _CURRENT_FILE_HANDLER = fh
        _CURRENT_LOG_PATH = target_path
    except Exception:
        pass


def close_account_log():
    global _CURRENT_FILE_HANDLER, _CURRENT_LOG_PATH
    root_logger = logging.getLogger()
    if _CURRENT_FILE_HANDLER is not None:
        try:
            _CURRENT_FILE_HANDLER.flush()
            _CURRENT_FILE_HANDLER.close()
        except Exception:
            pass
        try:
            root_logger.removeHandler(_CURRENT_FILE_HANDLER)
        except Exception:
            pass
        _CURRENT_FILE_HANDLER = None
        _CURRENT_LOG_PATH = None


def save_account_screenshot(page, account_name: str = "", label: str = "capture", max_keep: int = 100) -> str | None:
    """アカウント別フォルダ screenshots/<アカウント名>/ にスクショを保存"""
    if page is None:
        return None
    try:
        safe_acc = sanitize_filename_part(account_name or "system")
        safe_lbl = sanitize_filename_part(label or "capture")
        out_dir = get_screenshot_dir(safe_acc)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        filename = f"{ts}_{safe_lbl}.png"
        out_path = os.path.join(out_dir, filename)
        page.screenshot(path=out_path, timeout=10000)
        logger.info(f"📸 スクショ保存: screenshots/{safe_acc}/{filename}")

        if max_keep > 0:
            try:
                files = sorted(glob.glob(os.path.join(out_dir, "*.png")), key=os.path.getmtime)
                for old_file in files[:-max_keep]:
                    try:
                        os.remove(old_file)
                    except Exception:
                        pass
            except Exception:
                pass
        return out_path
    except Exception as e:
        logger.debug(f"スクショ保存スキップ: {e}")
        return None


switch_account_log(None)

LEGACY_IGNORED_SHEETS = {"設定", "日常ツイート"}
def normalize_text(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    if isinstance(val, dt_time):
        return val.strftime("%H:%M")
    return unicodedata.normalize("NFKC", str(val)).strip()


# ==========================================
# 凍結アカウントの即時永久除外メモリセット（再起動不要）
# ==========================================
BLACKLISTED_ACCOUNTS = set()
SUSPENDED_FILE = "suspended_accounts.txt"


def load_blacklist():
    global BLACKLISTED_ACCOUNTS
    if os.path.exists(SUSPENDED_FILE):
        try:
            with open(SUSPENDED_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    acc = line.split(",")[0].strip().replace("@", "").lower()
                    if acc:
                        BLACKLISTED_ACCOUNTS.add(acc)
            logger.info(f"🚫 登録済みブラックリスト: {len(BLACKLISTED_ACCOUNTS)} 件のアカウントを読み込みました")
        except Exception:
            pass


load_blacklist()

# ==========================================
# 【案3】画像クールダウン＆アカウント間重複防止ロック
# ==========================================
IMAGE_HISTORY_FILE = "image_history.json"
IMAGE_COOLDOWN_SAME_ACCOUNT_DAYS = 30  # 同一アカウント：30日間再選出禁止
IMAGE_COOLDOWN_CROSS_ACCOUNT_DAYS = 7  # 他アカウント：7日間共有禁止


def load_image_history():
    if os.path.exists(IMAGE_HISTORY_FILE):
        try:
            with open(IMAGE_HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_image_history(history):
    try:
        with open(IMAGE_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"画像履歴保存エラー: {e}")


def record_image_usage(image_path, handle):
    if not image_path:
        return
    history = load_image_history()
    img_name = os.path.basename(image_path)
    clean_handle = handle.replace("@", "").lower()
    if img_name not in history:
        history[img_name] = []
    history[img_name].append({
        "account": clean_handle,
        "used_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })
    save_image_history(history)
    logger.info(f"📸 [画像履歴更新] [{handle}] が「{img_name}」の使用を記録しました。")


def is_image_available(img_name, current_handle, history):
    clean_handle = current_handle.replace("@", "").lower()
    records = history.get(img_name, [])
    now = datetime.now()
    for rec in records:
        try:
            dt = datetime.strptime(rec["used_at"], "%Y-%m-%d %H:%M:%S")
            diff_days = (now - dt).total_seconds() / 86400.0
            rec_acc = rec.get("account", "").lower()
            if rec_acc == clean_handle and diff_days < IMAGE_COOLDOWN_SAME_ACCOUNT_DAYS:
                return False, f"同一アカウント利用から{int(diff_days)}日（30日待機中）"
            if rec_acc != clean_handle and diff_days < IMAGE_COOLDOWN_CROSS_ACCOUNT_DAYS:
                return False, f"他アカウント(@{rec_acc})利用から{int(diff_days)}日（7日待機中）"
        except Exception:
            pass
    return True, "OK"


def resolve_random_image_from_folder(raw_target, handle=""):
    target_path = str(raw_target or "").strip().replace('"', '').replace("'", "")
    if not target_path or target_path in ["(なし)", "なし", "none", "None"]:
        return None

    app_dir = get_app_dir()
    search_dirs = [
        target_path,
        os.path.join("images", target_path),
        os.path.join(app_dir, target_path),
        os.path.join(app_dir, "images", target_path),
    ]

    if os.path.isfile(target_path):
        return os.path.abspath(target_path)

    valid_folder = None
    for d in search_dirs:
        if d and os.path.isdir(d):
            valid_folder = os.path.abspath(d)
            break

    if not valid_folder:
        return None

    img_exts = ("*.jpg", "*.jpeg", "*.png", "*.webp", "*.gif")
    candidates = []
    for ext in img_exts:
        candidates.extend(glob.glob(os.path.join(valid_folder, ext)))
        candidates.extend(glob.glob(os.path.join(valid_folder, ext.upper())))

    if not candidates:
        return None

    candidates = list(set(candidates))
    history = load_image_history()

    available_files = []
    for fpath in candidates:
        fname = os.path.basename(fpath)
        ok, _ = is_image_available(fname, handle, history)
        if ok:
            available_files.append(fpath)

    if available_files:
        chosen = random.choice(available_files)
        logger.info(
            f"📸 画像フォルダ [{os.path.basename(valid_folder)}] から選出"
            f"（クールダウン合格: {os.path.basename(chosen)} / 候補数: {len(available_files)}/{len(candidates)}）"
        )
        return os.path.abspath(chosen)

    logger.warning(f"⚠️ [{os.path.basename(valid_folder)}] の全画像がクールダウン待機中です。最も過去に使用された画像を安全に再利用します。")

    def get_last_used_dt(fpath):
        fname = os.path.basename(fpath)
        records = history.get(fname, [])
        if not records:
            return datetime.min
        try:
            return max(datetime.strptime(r["used_at"], "%Y-%m-%d %H:%M:%S") for r in records)
        except Exception:
            return datetime.min

    oldest_file = min(candidates, key=get_last_used_dt)
    return os.path.abspath(oldest_file)


# ==========================================
# スピンタックス & 動的タグ置換
# ==========================================
def expand_spintax(text: str) -> str:
    pattern = re.compile(r'\{([^{}]+)\}')
    while True:
        match = pattern.search(text)
        if not match:
            break
        choices = match.group(1).split('|')
        chosen = random.choice(choices)
        text = text[:match.start()] + chosen + text[match.end():]
    return text


def replace_dynamic_tags(text: str) -> str:
    now = datetime.now()
    weekdays_jp = ["月", "火", "水", "木", "金", "土", "日"]
    weekday_str = weekdays_jp[now.weekday()]

    text = text.replace("{日付}", now.strftime("%Y年%m月%d日"))
    text = text.replace("{月日}", f"{now.month}月{now.day}日")
    text = text.replace("{曜日}", weekday_str)
    text = text.replace("{時間}", now.strftime("%H:%M"))
    text = text.replace("{時刻}", now.strftime("%H:%M"))
    text = text.replace("{午前午後}", "午前" if now.hour < 12 else "午後")
    return text


# ==========================================
# 時間帯・曜日・間隔パーサー（全角・日本語・複数範囲・日またぎ対応）
# ==========================================
def parse_single_time_to_minutes(token: str) -> int | None:
    """
    単一時刻文字列（例: "7", "07:00", "7時", "7時半", "7時30分", "22:00", "24:00"）を分(0..1440)へ変換
    """
    s = normalize_text(token)
    if not s:
        return None
    s = re.sub(r'\s+', '', s)
    # 先頭の不要な修飾語（例: 翌, 午前）を取り除く補助（ただし午後対応）
    is_pm = False
    if "午後" in s or "pm" in s.lower():
        is_pm = True
    s_clean = re.sub(r'(翌|午前|午後|am|pm)', '', s, flags=re.IGNORECASE)

    # パターン1: HH:MM または HH時MM分 / HH時半 / HH時
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
    """
    柔軟な投稿時間・実行時間判定:
      - 空白 / None / "0:00-0:00" / "24時間" -> True（常に投稿可能）
      - "7～10", "７～１０", "7〜10", "7-10", "7時から10時", "7時～10時", "7:00-10:00"
        -> 7:00 <= 現在時刻 <= 10:00 の間 True（6:59 や 10:01 は False）
      - 日またぎ: "22～2" -> 22:00〜翌2:00 (現在 >= 22:00 または 現在 <= 2:00)
      - 複数時間帯: "7～10, 18～21"（カンマ・読点区切り）-> いずれかに一致で True
    """
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
        # 区切り文字を "~" に統一（"から", "〜", "～", "−", "–", "—", "ー", "-", "to"）
        unified = re.sub(r'(?:から|〜|～|~|−|–|—|ー|-|\bto\b)', '~', seg, flags=re.IGNORECASE)
        parts = [p.strip() for p in unified.split('~') if p.strip()]
        if len(parts) < 2:
            # 単一時刻のみ書かれている場合（例: "8時" -> 8:00〜8:59相当ではなく一致判定、解釈不能なら次へ）
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
        # 0:00-0:00 は24時間稼働扱い
        if s_min == 0 and e_min == 0:
            return True

        if s_min <= e_min:
            # 通常帯（例: 7:00〜10:00 は 7:00〜10:00 ちょうどまで対象、10:01から対象外）
            if s_min <= cur_min <= e_min:
                return True
        else:
            # 日またぎ（例: 22:00〜翌2:00）
            if cur_min >= s_min or cur_min <= e_min:
                return True

    # 有効な書式が1つも解釈できない入力だった場合は安全のためブロックせずFalseにはしない？
    # いや、誤記でも意図せず発火しないよう valid_range_found なら False、未記載扱いなら True
    return False if valid_range_found else True


def is_today_allowed(weekday_spec, now_dt: datetime | None = None) -> bool:
    """
    曜日設定を判定:
      - 空白 / 毎日 / 全日 / * / all -> True
      - 平日 / 月～金 / 月-金 -> 月〜金 True
      - 土日 / 週末 -> 土日 True
      - 月水金 / 月,火,木 など -> 該当曜日のみ True
    """
    raw = normalize_text(weekday_spec)
    if not raw or raw.lower() in ("毎日", "全日", "すべて", "all", "*", "none", "nan"):
        return True

    if now_dt is None:
        now_dt = datetime.now()
    wd = now_dt.weekday()  # 0=月 .. 6=日

    if "平日" in raw or "月~金" in raw or "月-金" in raw or "月〜金" in raw:
        return wd < 5
    if "土日" in raw or "週末" in raw:
        return wd >= 5

    jp_days = ["月", "火", "水", "木", "金", "土", "日"]
    en_days = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

    matched_any_token = False
    for ch in jp_days:
        if ch in raw:
            matched_any_token = True
    low = raw.lower()
    for ed in en_days:
        if ed in low:
            matched_any_token = True

    if not matched_any_token:
        return True

    return (jp_days[wd] in raw) or (en_days[wd] in low)


def parse_interval_seconds(interval_spec) -> tuple[int, int]:
    """
    ポスト間隔（例: "120~360分", "60-120", "30秒", "1~2時間"）を (min_sec, max_sec) に変換
    """
    default_min, default_max = 120 * 60, 360 * 60
    raw = normalize_text(interval_spec)
    if not raw:
        return default_min, default_max

    unit_mult = 60
    if "秒" in raw or raw.lower().endswith("s") or "sec" in raw.lower():
        unit_mult = 1
    elif "時間" in raw or "hour" in raw.lower() or raw.lower().endswith("h"):
        unit_mult = 3600

    nums = re.findall(r'\d+(?:\.\d+)?', raw)
    if not nums:
        return default_min, default_max

    try:
        v1 = max(1, int(float(nums[0]) * unit_mult))
        v2 = max(v1, int(float(nums[1]) * unit_mult)) if len(nums) >= 2 else v1
        return min(v1, v2), max(v1, v2)
    except Exception:
        return default_min, default_max


def parse_post_mode(mode_spec) -> int:
    """0: 上から順 / 1: ランダム"""
    raw = normalize_text(mode_spec)
    if not raw:
        return 1
    if raw.startswith("0") or ("上" in raw and "ランダム" not in raw) or ("順" in raw and "ランダム" not in raw):
        return 0
    return 1


def parse_theme_active(active_spec) -> bool:
    raw = normalize_text(active_spec).lower()
    if not raw:
        return True
    if raw in ("無効", "停止", "off", "false", "0", "disable", "disabled", "×", "✕"):
        return False
    return True


ALL_EMOJIS = [
    "😀", "😃", "😄", "😁", "😆", "😅", "😂", "🤣", "😊", "😇", "🙂", "🙃", "😉", "😌", "😍", "🥰",
    "👍", "👏", "🙌", "👐", "🤲", "🤝", "🙏", "✌️", "🤞", "🤟", "🤘", "👌", "👈", "👉",
    "✨", "🌟", "⭐", "💫", "⚡", "💥", "🔥", "🌈", "☀️", "🌤️", "⛅", "🌥️", "☁️", "🌦️", "🌧️",
    "🎉", "🎊", "🎈", "🎁", "🎀", "🪄", "🎯", "🏆", "🥇", "🥈", "🥉", "🏅", "🎖️", "🎪", "🎨"
]
ALL_EMOJIS = list(dict.fromkeys(ALL_EMOJIS))
LAST_USED_EMOJI = None


def get_unique_random_emoji():
    global LAST_USED_EMOJI
    candidates = [e for e in ALL_EMOJIS if e != LAST_USED_EMOJI]
    chosen = random.choice(candidates)
    LAST_USED_EMOJI = chosen
    return chosen


def human_mouse_move(page, target_x=None, target_y=None):
    try:
        if target_x is None or target_y is None:
            target_x = random.randint(300, 800)
            target_y = random.randint(200, 600)
        steps = random.randint(3, 6)
        page.mouse.move(target_x, target_y, steps=steps)
        time.sleep(random.uniform(0.1, 0.25))
    except Exception:
        pass


def human_typing(page, text):
    for char in text:
        try:
            page.keyboard.type(char)
            if char in ["、", "。", "！", "？", "\n", " "]:
                time.sleep(random.uniform(0.18, 0.4))
            else:
                time.sleep(random.uniform(0.02, 0.06))
        except Exception:
            pass


def resolve_image_path(raw_path, handle=""):
    if not raw_path:
        return None
    clean = str(raw_path).replace('"', '').replace("'", "").strip()
    if not clean or clean in ["(なし)", "なし", "none", "None"]:
        return None
    if os.path.isdir(clean) or os.path.isdir(os.path.join("images", clean)):
        return resolve_random_image_from_folder(clean, handle=handle)
    if os.path.exists(clean):
        return os.path.abspath(clean)
    fname = os.path.basename(clean)
    candidate_dirs = [
        "images",
        os.path.join(get_app_dir(), "images"),
        os.path.abspath("images"),
        os.path.abspath(os.path.join("..", "images"))
    ]
    for cdir in candidate_dirs:
        test_path = os.path.join(cdir, fname)
        if os.path.exists(test_path):
            return os.path.abspath(test_path)
    return None


def find_excel_for_account(user_id):
    clean_id = user_id.replace("@", "").strip()
    candidates = [f"{clean_id}.xlsx", f"@{clean_id}.xlsx", f"tweets_{clean_id}.xlsx"]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


PAUSED_ACCOUNTS_FILE = "paused_accounts.json"


def load_paused_accounts() -> set[str]:
    p = os.path.join(get_app_dir(), PAUSED_ACCOUNTS_FILE)
    if not os.path.exists(p):
        return set()
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return {str(x).replace("@", "").strip().lower() for x in data if str(x).strip()}
    except Exception:
        pass
    return set()


def load_accounts(include_paused: bool = False):
    accounts_file = os.path.join(get_app_dir(), "アカウント情報.xlsx")
    if not os.path.exists(accounts_file):
        accounts_file = "アカウント情報.xlsx"
    accounts = []
    if not os.path.exists(accounts_file):
        return accounts
    paused_set = set() if include_paused else load_paused_accounts()
    try:
        wb = openpyxl.load_workbook(accounts_file, data_only=True)
        ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
        for row in range(2, ws.max_row + 1):
            user_id = str(ws.cell(row, 1).value or "").strip()
            pw = str(ws.cell(row, 2).value or "").strip()
            name = str(ws.cell(row, 3).value or "").strip()
            token = str(ws.cell(row, 4).value or "").strip()
            if not user_id:
                continue
            clean_id = user_id.replace("@", "").lower()

            if clean_id in BLACKLISTED_ACCOUNTS:
                continue
            if not include_paused and clean_id in paused_set:
                continue

            handle = user_id if user_id.startswith("@") else "@" + user_id
            excel_file = find_excel_for_account(user_id)
            accounts.append({
                "handle": handle,
                "user_id": user_id,
                "clean_id": clean_id,
                "name": name or handle,
                "password": pw,
                "auth_token": token,
                "excel_path": excel_file
            })
    except Exception as e:
        logger.error(f"アカウント情報読み込みエラー: {e}")
    return accounts


def blacklist_and_remove_account(target_handle, reason="凍結・制限検知"):
    global BLACKLISTED_ACCOUNTS
    clean_target = target_handle.replace("@", "").strip().lower()

    BLACKLISTED_ACCOUNTS.add(clean_target)
    logger.critical(f"🚫【即時除外】[{target_handle}] をメモリ上の永久ブロックリストに追加しました（アプリ再起動不要）")

    try:
        with open(SUSPENDED_FILE, "a", encoding="utf-8") as f:
            f.write(f"{target_handle},{datetime.now().strftime('%Y-%m-%d %H:%M:%S')},{reason}\n")
    except Exception:
        pass

    accounts_file = "アカウント情報.xlsx"
    if os.path.exists(accounts_file):
        for attempt in range(3):
            try:
                wb = openpyxl.load_workbook(accounts_file)
                ws = wb["アカウント情報"] if "アカウント情報" in wb.sheetnames else wb.active
                row_to_delete = None
                for r in range(2, ws.max_row + 1):
                    val = str(ws.cell(r, 1).value or "").strip().replace("@", "").lower()
                    if val == clean_target:
                        row_to_delete = r
                        break
                if row_to_delete:
                    ws.delete_rows(row_to_delete)
                    wb.save(accounts_file)
                    logger.info(f"🗑️「アカウント情報.xlsx」から [{target_handle}] の行を削除しました。")
                break
            except PermissionError:
                time.sleep(2)
            except Exception as e:
                logger.error(f"Excel削除エラー: {e}")
                break


def check_suspension_or_lock(page):
    suspension_keywords = [
        "アカウントが凍結されています", "アカウントは凍結されています", "アカウントがロックされています",
        "ご利用のアカウントは現在制限されています", "アカウントは一時的に制限されています",
        "Your account has been suspended", "Account suspended", "Your account is locked",
        "account is temporarily restricted", "アカウントへのアクセスが制限されています"
    ]
    try:
        content = page.content()
        for kw in suspension_keywords:
            if kw in content:
                return kw
    except Exception:
        pass
    return None


# ==========================================
# テーマ別マルチシート読み込み＆ステータス更新
# ==========================================
def _detect_sheet_layout(ws):
    """
    テーマシートのヘッダー構造（1〜2行目テーマ設定 + 4行目テーブルヘッダー）
    および旧シート（1行目テーブルヘッダー）の両方を判別する
    """
    a1 = normalize_text(ws.cell(1, 1).value)
    b1 = normalize_text(ws.cell(1, 2).value)
    c1 = normalize_text(ws.cell(1, 3).value)
    has_theme_banner = ("実行時間" in a1) and ("曜日" in b1 or "間隔" in c1)

    settings = {
        "time_range": "0:00-0:00",
        "weekdays": "毎日",
        "min_interval": 120 * 60,
        "max_interval": 360 * 60,
        "mode": 1,
        "enabled": True,
    }

    if has_theme_banner:
        settings["time_range"] = normalize_text(ws.cell(2, 1).value) or "0:00-0:00"
        settings["weekdays"] = normalize_text(ws.cell(2, 2).value) or "毎日"
        iv_min, iv_max = parse_interval_seconds(ws.cell(2, 3).value)
        settings["min_interval"] = iv_min
        settings["max_interval"] = iv_max
        settings["mode"] = parse_post_mode(ws.cell(2, 4).value)
        settings["enabled"] = parse_theme_active(ws.cell(2, 5).value)

    header_row = None
    scan_start = 3 if has_theme_banner else 1
    scan_end = min(15, max(ws.max_row, 4))
    for r in range(scan_start, scan_end + 1):
        vals = [normalize_text(ws.cell(r, c).value).lower() for c in range(1, min(ws.max_column, 12) + 1)]
        if any(v in ("ツイート文章", "ポスト文章", "ツイート本文", "ポスト本文", "本文", "text") for v in vals):
            header_row = r
            break

    if header_row is None:
        header_row = 4 if has_theme_banner else 1

    text_col = 2
    post_time_col = 3 if has_theme_banner else None
    img_col = 4 if has_theme_banner else 3
    status_col = 5 if has_theme_banner else 4
    date_col = 6 if has_theme_banner else 5

    for c in range(1, ws.max_column + 1):
        h = normalize_text(ws.cell(header_row, c).value).lower()
        if h in ("ツイート文章", "ポスト文章", "ツイート本文", "ポスト本文", "本文", "text"):
            text_col = c
        elif h in ("投稿時間", "投稿時間帯", "時間指定", "post_time", "time"):
            post_time_col = c
        elif h in ("画像パス", "画像", "画像フォルダ", "image_path", "image"):
            img_col = c
        elif h in ("ステータス", "status"):
            status_col = c
        elif h in ("投稿日時", "posted_at"):
            date_col = c

    return {
        "has_theme_banner": has_theme_banner,
        "settings": settings,
        "header_row": header_row,
        "text_col": text_col,
        "post_time_col": post_time_col,
        "img_col": img_col,
        "status_col": status_col,
        "date_col": date_col,
    }


def reset_all_tweets_status(excel_path: str, sheet_name: str, header_row: int, text_col: int, status_col: int, date_col: int | None = None):
    try:
        wb = openpyxl.load_workbook(excel_path)
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        for r in range(header_row + 1, ws.max_row + 1):
            if ws.cell(r, text_col).value:
                st_cell = ws.cell(r, status_col)
                st_cell.value = ""
                st_cell.fill = PatternFill(
                    start_color="F1F5F9" if r % 2 == 0 else "FFFFFF",
                    end_color="F1F5F9" if r % 2 == 0 else "FFFFFF",
                    fill_type="solid"
                )
                if date_col:
                    ws.cell(r, date_col).value = ""
        wb.save(excel_path)
        logger.info(f"[{os.path.basename(excel_path)} / {sheet_name}] 全ポスト完了のためステータスを自動リセットしました。")
    except Exception as e:
        logger.warning(f"ステータスリセットエラー ({sheet_name}): {e}")


def load_theme_sheets(excel_path: str, handle: str = "", now_dt: datetime | None = None) -> list[dict]:
    """
    1つのアカウントExcel内の全テーマシートを読み込む。
    返却リストの各要素は 1つのテーマシートを表し、
    「投稿時間」列でフィルタ済みの即投稿可能ツイート(eligible_tweets)と未投稿全体件数を含む。
    """
    themes = []
    if not excel_path or not os.path.exists(excel_path):
        return themes
    if now_dt is None:
        now_dt = datetime.now()

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        for sheet_name in wb.sheetnames:
            if sheet_name in LEGACY_IGNORED_SHEETS:
                continue
            ws = wb[sheet_name]
            layout = _detect_sheet_layout(ws)
            st = layout["settings"]

            header_row = layout["header_row"]
            text_col = layout["text_col"]
            post_time_col = layout["post_time_col"]
            img_col = layout["img_col"]
            status_col = layout["status_col"]
            date_col = layout["date_col"]

            total_count = 0
            unposted_tweets = []
            eligible_tweets = []
            skipped_by_time_count = 0

            for r in range(header_row + 1, ws.max_row + 1):
                raw_text = ws.cell(r, text_col).value
                text = str(raw_text).strip() if raw_text is not None else ""
                if not text:
                    continue
                total_count += 1

                status = normalize_text(ws.cell(r, status_col).value).lower()
                if status in ("posted", "投稿済み", "投稿済", "完了", "skipped", "スキップ"):
                    continue

                raw_pt = ws.cell(r, post_time_col).value if post_time_col else ""
                pt_str = normalize_text(raw_pt)
                raw_img = ws.cell(r, img_col).value if img_col else ""

                item = {
                    "row": r,
                    "raw_text": text,
                    "post_time": pt_str,
                    "raw_img": str(raw_img or "").strip(),
                    "status_col": status_col,
                    "date_col": date_col,
                }
                unposted_tweets.append(item)

                # 投稿時間判定：未設定(空白)なら通常通り対象、設定あり＆時間外ならステータスを変えず即スキップ！
                if pt_str and not is_time_in_ranges(pt_str, now_dt=now_dt):
                    skipped_by_time_count += 1
                    continue

                eligible_tweets.append(item)

            themes.append({
                "sheet_name": sheet_name,
                "settings": st,
                "layout": layout,
                "total_count": total_count,
                "unposted_count": len(unposted_tweets),
                "skipped_by_time_count": skipped_by_time_count,
                "eligible_tweets": eligible_tweets,
            })
    except Exception as e:
        logger.error(f"テーマ読み込みエラー ({os.path.basename(excel_path)}): {e}")
    return themes


def update_excel_status(excel_path: str, sheet_name: str, row_idx: int, status_col: int, date_col: int, status_text: str):
    for attempt in range(3):
        try:
            wb = openpyxl.load_workbook(excel_path)
            if sheet_name not in wb.sheetnames:
                return
            ws = wb[sheet_name]
            st_cell = ws.cell(row_idx, status_col)
            st_cell.value = status_text

            # ステータス表示を見やすい色で維持（削除耐性・条件付き書式とも協調）
            st_low = status_text.lower()
            if st_low in ("posted", "投稿済み", "投稿済", "完了"):
                st_cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                st_cell.font = Font(name="Yu Gothic UI", size=10, bold=True, color="15803D")
            elif st_low in ("skipped", "スキップ"):
                st_cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
                st_cell.font = Font(name="Yu Gothic UI", size=10, bold=False, color="475569")
            elif st_low in ("failed", "失敗", "エラー"):
                st_cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                st_cell.font = Font(name="Yu Gothic UI", size=10, bold=True, color="B91C1C")

            ws.cell(row_idx, date_col).value = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            wb.save(excel_path)
            logger.info(f"Excel更新完了: [{sheet_name}] 行{row_idx} -> {status_text}")
            return
        except PermissionError:
            time.sleep(2)
        except Exception as e:
            logger.warning(f"Excel更新エラー: {e}")
            break


# 画像・動画ブロック（Twitter/Xへのアップロード通信は絶対に遮断しない）
def apply_media_blocking(page):
    def route_filter(route):
        req = route.request
        if "upload.twitter.com" in req.url or "upload.x.com" in req.url or "blob:" in req.url or req.method != "GET":
            return route.continue_()

        if req.resource_type in ["image", "media", "font"]:
            return route.abort()
        url_clean = req.url.lower().split("?")[0]
        if any(url_clean.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".gif", ".webp", ".mp4", ".webm", ".woff", ".woff2", ".ttf"]):
            return route.abort()
        return route.continue_()
    try:
        page.route("**/*", route_filter)
        logger.info("⚡ メディア通信ブロック稼働中 (画像アップロード通信は正常許可)")
    except Exception:
        pass


def check_for_errors_or_duplicates(page):
    duplicate_texts = ["すでに同じツイートを送信済みです", "重複したポストです", "You have already sent this"]
    error_texts = ["エラーが発生しました", "送信できませんでした", "Something went wrong"]
    try:
        content = page.content()
        for dup in duplicate_texts:
            if dup in content:
                return "duplicate"
        for err in error_texts:
            if err in content:
                return "general_error"
    except Exception:
        pass
    return None


def execute_single_post_attempt(page, text_to_send, image_path, account_name: str = ""):
    try:
        if is_stop_requested():
            return "aborted"
        page.goto("https://x.com/compose/post", wait_until="domcontentloaded", timeout=30000)
        if interruptible_sleep(random.uniform(2.0, 3.0)):
            return "aborted"

        susp_msg = check_suspension_or_lock(page)
        if susp_msg:
            save_account_screenshot(page, account_name, "suspended")
            return f"suspended:{susp_msg}"

        human_mouse_move(page)
        editor = page.locator('div[role="textbox"]').first
        try:
            editor.wait_for(state="visible", timeout=8000)
        except Exception:
            if is_stop_requested():
                return "aborted"
            susp_msg = check_suspension_or_lock(page)
            if susp_msg:
                save_account_screenshot(page, account_name, "suspended")
                return f"suspended:{susp_msg}"
            save_account_screenshot(page, account_name, "error_textbox_missing")
            return "general_error"

        if is_stop_requested():
            return "aborted"

        box = editor.bounding_box()
        if box:
            page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2, steps=4)
            time.sleep(0.1)
            page.mouse.down()
            time.sleep(0.05)
            page.mouse.up()
        else:
            editor.click()

        if interruptible_sleep(random.uniform(0.3, 0.6)):
            return "aborted"
        human_typing(page, text_to_send)
        if interruptible_sleep(random.uniform(0.8, 1.5)):
            return "aborted"

        if image_path and os.path.exists(image_path):
            try:
                logger.info(f"📸 画像ファイルを添付中: {os.path.basename(image_path)}")
                file_input = page.locator("input[type='file'][data-testid='fileInput']").first
                file_input.set_input_files(image_path)
                if interruptible_sleep(random.uniform(3.5, 5.0)):
                    return "aborted"
            except Exception as img_err:
                logger.warning(f"画像添付エラー: {img_err}")

        if interruptible_sleep(random.uniform(0.5, 1.0)):
            return "aborted"

        btn = page.locator('[data-testid="tweetButton"]').first
        btn_box = btn.bounding_box()
        if btn_box:
            page.mouse.move(btn_box["x"] + btn_box["width"] / 2, btn_box["y"] + btn_box["height"] / 2, steps=5)
            time.sleep(random.uniform(0.2, 0.4))
            btn.click()
        else:
            page.keyboard.press("Control+Enter")

        time.sleep(random.uniform(3.0, 4.0))
        susp_msg = check_suspension_or_lock(page)
        if susp_msg:
            save_account_screenshot(page, account_name, "suspended")
            return f"suspended:{susp_msg}"

        err_type = check_for_errors_or_duplicates(page)
        if err_type:
            save_account_screenshot(page, account_name, f"error_{err_type}")
            return err_type
        save_account_screenshot(page, account_name, "posted")
        return "success"
    except Exception as e:
        if is_stop_requested():
            return "aborted"
        logger.error(f"投稿実行エラー: {e}")
        save_account_screenshot(page, account_name, "error_exception")
        return "general_error"


def post_tweet_with_smart_retry(page, original_text, image_path, account_name: str = ""):
    res1 = execute_single_post_attempt(page, original_text, image_path, account_name=account_name)
    if res1 == "aborted" or is_stop_requested():
        return "aborted"
    if res1.startswith("suspended:"):
        return res1
    if res1 == "success":
        logger.info("=== ポスト送信完了 ===")
        return "success"

    chosen_emoji = get_unique_random_emoji()
    spaced_text = f"{original_text} {chosen_emoji}"
    logger.warning(f"【送信失敗/重複: {res1}】絵文字「 {chosen_emoji} 」を付与してリトライ...")
    try:
        page.keyboard.press("Escape")
        time.sleep(1)
    except Exception:
        pass

    if is_stop_requested():
        return "aborted"

    res2 = execute_single_post_attempt(page, spaced_text, image_path, account_name=account_name)
    if res2 == "aborted" or is_stop_requested():
        return "aborted"
    if res2.startswith("suspended:"):
        return res2
    if res2 == "success":
        logger.info(f"=== リトライによりポスト送信完了（絵文字: {chosen_emoji}） ===")
        return "success"

    logger.error("【リトライも失敗】スキップして次へ進みます。")
    try:
        page.keyboard.press("Escape")
        time.sleep(1)
    except Exception:
        pass
    return "skipped"


def resolve_session_dir(handle: str) -> str:
    """chrome_launcher.exe と共通のルールでプロファイル保存先を解決"""
    app_sessions = os.path.join(get_app_dir(), "browser_sessions")
    cwd_sessions = os.path.abspath("browser_sessions")
    base_dir = app_sessions if os.path.exists(app_sessions) or not os.path.exists(cwd_sessions) else cwd_sessions
    os.makedirs(base_dir, exist_ok=True)

    clean_id = re.sub(r'[<>:"/\\|?*]', "_", handle.replace("@", "").strip())
    exact_dir = os.path.join(base_dir, clean_id)
    underscore_dir = os.path.join(base_dir, f"_{clean_id}")

    if os.path.exists(exact_dir):
        return os.path.abspath(exact_dir)
    if os.path.exists(underscore_dir):
        return os.path.abspath(underscore_dir)

    os.makedirs(exact_dir, exist_ok=True)
    return os.path.abspath(exact_dir)


def launch_browser_safely(p, session_dir):
    kwargs = {
        "user_data_dir": session_dir,
        "headless": False,
        "viewport": {"width": 1280, "height": 800},
        "args": ["--disable-blink-features=AutomationControlled"]
    }
    try:
        ctx = p.chromium.launch_persistent_context(**kwargs, channel="chrome")
        logger.info("🌐 [ブラウザ起動] Google Chrome を使用します。")
        return ctx
    except Exception:
        pass
    try:
        ctx = p.chromium.launch_persistent_context(**kwargs, channel="msedge")
        logger.info("🌐 [ブラウザ起動] Microsoft Edge を使用します。")
        return ctx
    except Exception:
        pass
    logger.info("🌐 [ブラウザ起動] 標準 Chromium を使用します。")
    return p.chromium.launch_persistent_context(**kwargs)


def _has_cmd_special_chars(p: str) -> bool:
    if not isinstance(p, str):
        return False
    return any(ch in p for ch in ("(", ")", "&", "^", " "))


def _to_short_path(path_str: str) -> str:
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(1024)
        ret = ctypes.windll.kernel32.GetShortPathNameW(path_str, buf, 1024)
        if ret and buf.value:
            return buf.value
    except Exception:
        pass
    return path_str


def _ensure_clean_driver_junction(real_driver_dir: str) -> str | None:
    if not os.path.isdir(real_driver_dir):
        return None

    candidate_parents = [
        os.environ.get("LOCALAPPDATA", ""),
        tempfile.gettempdir(),
        os.environ.get("PUBLIC", r"C:\Users\Public"),
        r"C:\ProgramData",
    ]
    for parent in candidate_parents:
        if not parent or not os.path.isdir(parent):
            continue
        if _has_cmd_special_chars(parent):
            continue
        junction_path = os.path.join(parent, "x_bot_pw_driver")
        try:
            if os.path.exists(junction_path):
                try:
                    same_target = (
                        os.path.realpath(junction_path).lower() == os.path.realpath(real_driver_dir).lower()
                    )
                    if same_target:
                        return junction_path
                except Exception:
                    pass
                subprocess.run(
                    ["cmd", "/c", "rmdir", junction_path],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )

            if not os.path.exists(junction_path):
                res = subprocess.run(
                    ["cmd", "/c", "mklink", "/J", junction_path, real_driver_dir],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                if res.returncode == 0 and os.path.isdir(junction_path):
                    return junction_path
        except Exception:
            continue
    return None


def _build_clean_cmd_wrapper(cmd_path: str) -> str:
    short_p = _to_short_path(cmd_path)
    if not _has_cmd_special_chars(short_p):
        return short_p

    driver_dir = os.path.dirname(cmd_path)
    node_exe = os.path.join(driver_dir, "node.exe")
    cli_js = os.path.join(driver_dir, "package", "cli.js")

    candidate_dirs = [
        tempfile.gettempdir(),
        os.environ.get("PUBLIC", r"C:\Users\Public"),
        r"C:\ProgramData",
    ]
    for cdir in candidate_dirs:
        if not cdir or not os.path.isdir(cdir):
            continue
        wrapper_path = os.path.join(cdir, "x_bot_pw_driver.cmd")
        if _has_cmd_special_chars(wrapper_path):
            continue
        try:
            if os.path.isfile(node_exe) and os.path.isfile(cli_js):
                body = (
                    "@echo off\r\n"
                    'if "%~nx1"=="cli.js" goto has_cli\r\n'
                    f'"{node_exe}" "{cli_js}" %*\r\n'
                    "exit /b %ERRORLEVEL%\r\n"
                    ":has_cli\r\n"
                    f'"{node_exe}" %*\r\n'
                    "exit /b %ERRORLEVEL%\r\n"
                )
            else:
                body = f'@echo off\r\ncall "{cmd_path}" %*\r\n'
            try:
                with open(wrapper_path, "w", encoding="mbcs") as wf:
                    wf.write(body)
            except Exception:
                with open(wrapper_path, "w", encoding="utf-8") as wf:
                    wf.write("@echo off\r\nchcp 65001 >nul\r\n" + body.replace("@echo off\r\n", "", 1))
            return wrapper_path
        except Exception:
            continue
    return cmd_path


def fix_playwright_windows_driver():
    if sys.platform != "win32":
        return

    try:
        import importlib
        for mod_path in (
            "playwright._impl._driver",
            "playwright._impl._transport",
            "playwright.sync_api._context_manager",
            "playwright.async_api._context_manager",
        ):
            try:
                importlib.import_module(mod_path)
            except Exception:
                pass

        import playwright._impl._driver as pw_driver
        orig_compute = pw_driver.compute_driver_executable

        raw_res = orig_compute()
        logger.info(f"🔍 Playwright検出ドライバ: {raw_res}")

        first_item = raw_res[0] if isinstance(raw_res, tuple) and raw_res else raw_res
        clean_driver_dir = None
        if isinstance(first_item, str) and _has_cmd_special_chars(first_item):
            real_driver_dir = os.path.dirname(first_item)
            clean_driver_dir = _ensure_clean_driver_junction(real_driver_dir)
            if clean_driver_dir:
                logger.info(f"🔧 Playwrightクリーンパス作成: {clean_driver_dir}")

        def _rewrite_path(p: str) -> str:
            if not isinstance(p, str):
                return p
            if clean_driver_dir:
                real_driver_dir = os.path.dirname(first_item)
                try:
                    rel = os.path.relpath(p, real_driver_dir)
                    if not rel.startswith(".."):
                        candidate = os.path.join(clean_driver_dir, rel)
                        if os.path.exists(candidate):
                            return candidate
                except Exception:
                    pass
            if p.lower().endswith((".cmd", ".bat")) and _has_cmd_special_chars(p):
                return _build_clean_cmd_wrapper(p)
            return p

        def _safe_compute_driver_executable():
            res = orig_compute()
            if isinstance(res, tuple):
                return tuple(_rewrite_path(x) for x in res)
            if isinstance(res, str):
                return _rewrite_path(res)
            return res

        patched_count = 0
        for mod_name, mod in list(sys.modules.items()):
            if mod_name.startswith("playwright") and mod is not None:
                if hasattr(mod, "compute_driver_executable"):
                    setattr(mod, "compute_driver_executable", _safe_compute_driver_executable)
                    patched_count += 1
        logger.info(f"🔧 Playwrightドライバ補正完了 ({patched_count} モジュール適用)")

        if not getattr(asyncio.create_subprocess_exec, "_xbot_patched", False):
            orig_cse = asyncio.create_subprocess_exec

            async def _safe_create_subprocess_exec(program, *args, **kwargs):
                if isinstance(program, str) and _has_cmd_special_chars(program):
                    new_prog = _rewrite_path(program)
                    if new_prog != program:
                        logger.info(f"🔧 サブプロセス経路補正: {os.path.basename(program)} -> {new_prog}")
                        program = new_prog
                new_args = tuple(_rewrite_path(a) if isinstance(a, str) else a for a in args)
                return await orig_cse(program, *new_args, **kwargs)

            _safe_create_subprocess_exec._xbot_patched = True
            asyncio.create_subprocess_exec = _safe_create_subprocess_exec

    except Exception as e:
        logger.warning(f"⚠️ Playwrightドライバパス補正エラー: {e}")


def get_pid_file() -> str:
    return os.path.join(get_app_dir(), "bot_running.pid")


def get_stop_flag_file() -> str:
    return os.path.join(get_app_dir(), "stop_bot.flag")


def get_active_profile_file() -> str:
    return os.path.join(get_app_dir(), "bot_active_profile.txt")


def set_active_bot_profile(profile_name: str | None) -> None:
    path = get_active_profile_file()
    try:
        if profile_name:
            with open(path, "w", encoding="utf-8") as f:
                f.write(sanitize_filename_part(profile_name))
        else:
            if os.path.exists(path):
                os.remove(path)
    except Exception:
        pass


def get_status_file() -> str:
    return os.path.join(get_app_dir(), "bot_status.json")


def write_bot_status(current_account: str = "", current_theme: str = "", next_post_times: dict | None = None) -> None:
    try:
        serialized_next = {}
        if next_post_times:
            for (cid, sname), dt_val in next_post_times.items():
                serialized_next[f"{cid}|{sname}"] = dt_val.strftime("%Y-%m-%d %H:%M:%S")
        payload = {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "current_account": current_account,
            "current_theme": current_theme,
            "next_post_times": serialized_next,
        }
        with open(get_status_file(), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def is_stop_requested() -> bool:
    return os.path.exists(get_stop_flag_file())


def interruptible_sleep(total_seconds: float) -> bool:
    """停止フラグを0.5秒ごとに監視しながら待機する。停止要求があれば True を返す"""
    end_at = time.time() + max(0.0, float(total_seconds))
    while time.time() < end_at:
        if is_stop_requested():
            return True
        time.sleep(min(0.5, max(0.05, end_at - time.time())))
    return is_stop_requested()


def safe_run():
    logger.info("==================================================")
    logger.info("★ X 自動操縦ボット起動 (テーマ別シート管理＆投稿時間フィルタ対応版)")
    logger.info("★ テーマ別管理: Excelの各シート(タブ)を独立テーマとして運用")
    logger.info("★ 投稿時間フィルタ: 時間外ポストは即スキップして次の投稿へ進みます")
    logger.info("★ 案3: 画像クールダウン（同一垢30日 / 他垢7日共有ロック）")
    logger.info("★ 1窓交代ローテーション ＆ 画像・動画通信ブロック(メモリ60%カット)")
    logger.info("★ 凍結対策: 凍結検知時にメモリから即座に永久排除 ＆ 再起動不要で継続！")
    logger.info(f"★ 登録済みブラックリスト除外アカウント数: {len(BLACKLISTED_ACCOUNTS)} 件")
    logger.info("==================================================")

    pid_path = get_pid_file()
    stop_flag_path = get_stop_flag_file()
    try:
        if os.path.exists(stop_flag_path):
            os.remove(stop_flag_path)
        with open(pid_path, "w", encoding="utf-8") as pf:
            pf.write(str(os.getpid()))
    except Exception:
        pass

    next_post_times = {}  # key: (clean_id, sheet_name)
    fix_playwright_windows_driver()

    try:
        with sync_playwright() as p:
            while not is_stop_requested():
                BLACKLISTED_ACCOUNTS.clear()
                load_blacklist()
                accounts = load_accounts()
                if not accounts:
                    write_bot_status("", "", next_post_times)
                    logger.warning("運用可能な健全アカウントが見つかりません（全アカウント停止中または未登録）。30秒後に再確認します。")
                    if interruptible_sleep(30):
                        break
                    continue

                worked_in_this_cycle = False

                for acc in accounts:
                    if is_stop_requested():
                        break

                    handle, clean_id = acc["handle"], acc["clean_id"]
                    excel_path = acc.get("excel_path")
                    if clean_id in BLACKLISTED_ACCOUNTS or not excel_path:
                        continue

                    switch_account_log(clean_id)
                    account_suspended = False
                    try:
                        now = datetime.now()
                        themes = load_theme_sheets(excel_path, handle=handle, now_dt=now)

                        for theme in themes:
                            if is_stop_requested():
                                break
                            if clean_id in BLACKLISTED_ACCOUNTS:
                                account_suspended = True
                                break

                            sheet_name = theme["sheet_name"]
                            st = theme["settings"]
                            layout = theme["layout"]

                            if not st["enabled"]:
                                continue
                            if not is_today_allowed(st["weekdays"], now_dt=now):
                                continue
                            if not is_time_in_ranges(st["time_range"], now_dt=now):
                                continue

                            schedule_key = (clean_id, sheet_name)
                            if schedule_key in next_post_times and datetime.now() < next_post_times[schedule_key]:
                                continue

                            eligible = theme["eligible_tweets"]
                            # このテーマの全ポストが投稿済みの場合のみステータスをリセット（投稿時間待ちがある間はリセットしない）
                            if theme["unposted_count"] == 0 and theme["total_count"] > 0:
                                reset_all_tweets_status(
                                    excel_path, sheet_name,
                                    layout["header_row"], layout["text_col"], layout["status_col"], layout["date_col"]
                                )
                                reloaded = load_theme_sheets(excel_path, handle=handle, now_dt=datetime.now())
                                matched = next((t for t in reloaded if t["sheet_name"] == sheet_name), None)
                                eligible = matched["eligible_tweets"] if matched else []
                            elif not eligible and theme["skipped_by_time_count"] > 0:
                                logger.info(
                                    f"⏭️ [{handle}] テーマ「{sheet_name}」: 未投稿 {theme['skipped_by_time_count']} 件は指定投稿時間外のため即スキップします。"
                                )
                                continue

                            if not eligible:
                                continue

                            if is_stop_requested():
                                break

                            target_tweet = random.choice(eligible) if st["mode"] == 1 else eligible[0]
                            target_text = replace_dynamic_tags(expand_spintax(target_tweet["raw_text"]))
                            target_img = resolve_image_path(target_tweet["raw_img"], handle=handle)

                            img_label = f" (画像: {os.path.basename(target_img)})" if target_img else " (画像なし)"
                            pt_label = f" [指定時間: {target_tweet['post_time']}]" if target_tweet["post_time"] else " [指定時間: なし]"
                            logger.info("\n--------------------------------------------------")
                            logger.info(f"🔄 [{handle}] テーマ「{sheet_name}」行{target_tweet['row']}{pt_label}{img_label}")
                            logger.info(f"📝 {target_text[:40]}...")
                            logger.info("--------------------------------------------------")

                            session_dir = resolve_session_dir(handle)
                            logger.info(f"📂 プロファイル参照先: {session_dir}")

                            context, page, is_suspended, suspension_reason = None, None, False, ""
                            set_active_bot_profile(clean_id)
                            write_bot_status(clean_id, sheet_name, next_post_times)
                            try:
                                context = launch_browser_safely(p, session_dir)
                                page = context.pages[0] if len(context.pages) > 0 else context.new_page()
                                apply_media_blocking(page)

                                if acc.get("auth_token"):
                                    context.add_cookies([{"name": "auth_token", "value": acc["auth_token"], "domain": ".x.com", "path": "/"}])

                                if is_stop_requested():
                                    break

                                page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=30000)
                                if interruptible_sleep(random.uniform(2.0, 3.0)):
                                    break

                                susp_check = check_suspension_or_lock(page)
                                if susp_check:
                                    is_suspended, suspension_reason = True, susp_check
                                    save_account_screenshot(page, clean_id, "suspended")
                                else:
                                    res = post_tweet_with_smart_retry(page, target_text, target_img, account_name=clean_id)
                                    if res == "aborted" or is_stop_requested():
                                        logger.info(f"🛑 [{handle}] 停止指示により投稿処理を中断しました（Excelステータスは未投稿のまま維持）。")
                                        break
                                    elif res.startswith("suspended:"):
                                        is_suspended, suspension_reason = True, res.replace("suspended:", "")
                                    elif res == "success":
                                        if target_img:
                                            record_image_usage(target_img, handle)
                                        update_excel_status(
                                            excel_path, sheet_name,
                                            target_tweet["row"], target_tweet["status_col"], target_tweet["date_col"],
                                            "posted"
                                        )
                                    else:
                                        update_excel_status(
                                            excel_path, sheet_name,
                                            target_tweet["row"], target_tweet["status_col"], target_tweet["date_col"],
                                            "skipped"
                                        )
                            except Exception as loop_e:
                                logger.error(f"[{handle}] ({sheet_name}) エラー: {loop_e}")
                                save_account_screenshot(page, clean_id, "loop_error")
                            finally:
                                if context:
                                    try:
                                        context.close()
                                        logger.info(f"🧹 [{handle}] の自動投稿用ブラウザを安全に終了しました。")
                                    except Exception:
                                        pass
                                set_active_bot_profile(None)

                            if is_suspended:
                                logger.critical(f"🚨 [{handle}] の凍結検知（原因: {suspension_reason}）")
                                blacklist_and_remove_account(handle, reason=suspension_reason)
                                for k in [key for key in next_post_times if key[0] == clean_id]:
                                    del next_post_times[k]
                                write_bot_status("", "", next_post_times)
                                logger.info("⚡ 再起動せず即時排除しました。残りのアカウントで継続します。")
                                if interruptible_sleep(2):
                                    break
                                account_suspended = True
                                break

                            t_min = min(st["min_interval"], st["max_interval"])
                            t_max = max(st["min_interval"], st["max_interval"])
                            wait_sec = random.randint(t_min, t_max) if t_min < t_max else t_min
                            next_post_times[schedule_key] = datetime.now() + timedelta(seconds=wait_sec)
                            write_bot_status("", "", next_post_times)

                            logger.info(
                                f"⏳ [{handle} / {sheet_name}] 次回: {next_post_times[schedule_key].strftime('%H:%M:%S')} "
                                f"({wait_sec // 60}分 {wait_sec % 60}秒 後)"
                            )
                            worked_in_this_cycle = True
                            if interruptible_sleep(random.uniform(5.0, 10.0)):
                                break
                    finally:
                        switch_account_log(None)

                    if account_suspended:
                        worked_in_this_cycle = True
                        continue

                if is_stop_requested():
                    break
                if not worked_in_this_cycle:
                    write_bot_status("", "", next_post_times)
                    if interruptible_sleep(15):
                        break

        if is_stop_requested():
            logger.info("🛑 停止指示を検知したため、ボットを安全に停止しました。")
    finally:
        write_bot_status("", "", next_post_times)
        for fp in (pid_path, stop_flag_path, get_active_profile_file()):
            try:
                if os.path.exists(fp):
                    os.remove(fp)
            except Exception:
                pass


def run_single_test(target_account: str, target_theme: str = ""):
    """GUIからの指定アカウント1件テスト投稿モード"""
    clean_target = sanitize_filename_part(target_account).lower()
    logger.info("==================================================")
    logger.info(f"⚡ 1件テスト投稿モード起動: 対象={clean_target}" + (f" / テーマ={target_theme}" if target_theme else ""))
    logger.info("==================================================")

    accounts = load_accounts(include_paused=True)
    acc = next((a for a in accounts if a["clean_id"].lower() == clean_target), None)
    if not acc:
        logger.error(f"❌ 対象アカウント「{clean_target}」が アカウント情報.xlsx に見つかりません（ブラックリスト除外済みの可能性あり）。")
        return

    excel_path = acc.get("excel_path")
    if not excel_path:
        logger.error(f"❌ [{acc['handle']}] に対応する Excel ファイルが見つかりません。")
        return

    pid_path = get_pid_file()
    stop_flag_path = get_stop_flag_file()
    try:
        if os.path.exists(stop_flag_path):
            os.remove(stop_flag_path)
        with open(pid_path, "w", encoding="utf-8") as pf:
            pf.write(str(os.getpid()))
    except Exception:
        pass

    fix_playwright_windows_driver()
    switch_account_log(acc["clean_id"])
    try:
        now = datetime.now()
        themes = load_theme_sheets(excel_path, handle=acc["handle"], now_dt=now)
        if target_theme:
            themes = [t for t in themes if t["sheet_name"] == target_theme]

        chosen_theme = None
        chosen_tweet = None
        for t in themes:
            if t["eligible_tweets"]:
                chosen_theme = t
                chosen_tweet = random.choice(t["eligible_tweets"]) if t["settings"]["mode"] == 1 else t["eligible_tweets"][0]
                break

        if not chosen_theme or not chosen_tweet:
            logger.warning(f"⚠️ [{acc['handle']}] テスト投稿可能な未投稿ポストがありません（投稿時間外または全件投稿済み）。")
            return

        target_text = replace_dynamic_tags(expand_spintax(chosen_tweet["raw_text"]))
        target_img = resolve_image_path(chosen_tweet["raw_img"], handle=acc["handle"])
        session_dir = resolve_session_dir(acc["handle"])

        logger.info(f"🧪 テスト対象: テーマ「{chosen_theme['sheet_name']}」行{chosen_tweet['row']} / {target_text[:40]}...")
        set_active_bot_profile(acc["clean_id"])
        with sync_playwright() as p:
            context, page = None, None
            try:
                context = launch_browser_safely(p, session_dir)
                page = context.pages[0] if len(context.pages) > 0 else context.new_page()
                apply_media_blocking(page)
                if acc.get("auth_token"):
                    context.add_cookies([{"name": "auth_token", "value": acc["auth_token"], "domain": ".x.com", "path": "/"}])

                page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=30000)
                if interruptible_sleep(random.uniform(2.0, 3.0)):
                    return

                susp = check_suspension_or_lock(page)
                if susp:
                    save_account_screenshot(page, acc["clean_id"], "suspended")
                    logger.critical(f"🚨 凍結・ロック検知: {susp}")
                    blacklist_and_remove_account(acc["handle"], reason=susp)
                else:
                    res = post_tweet_with_smart_retry(page, target_text, target_img, account_name=acc["clean_id"])
                    if res == "success":
                        if target_img:
                            record_image_usage(target_img, acc["handle"])
                        update_excel_status(
                            excel_path, chosen_theme["sheet_name"],
                            chosen_tweet["row"], chosen_tweet["status_col"], chosen_tweet["date_col"],
                            "posted"
                        )
                        logger.info("✅ 1件テスト投稿が完了しました！")
                    elif res == "aborted":
                        logger.info("🛑 テスト投稿を中断しました。")
                    elif res.startswith("suspended:"):
                        blacklist_and_remove_account(acc["handle"], reason=res.replace("suspended:", ""))
                    else:
                        logger.warning(f"⚠️ テスト結果: {res}")
            finally:
                if context:
                    try:
                        context.close()
                    except Exception:
                        pass
                set_active_bot_profile(None)
    finally:
        switch_account_log(None)
        for fp in (pid_path, stop_flag_path, get_active_profile_file()):
            try:
                if os.path.exists(fp):
                    os.remove(fp)
            except Exception:
                pass


def run_session_check():
    """全プロファイルの X ログイン状態を軽量ヘッドレスチェックし session_status.json へ保存"""
    logger.info("==================================================")
    logger.info("🔍 全プロファイル X ログイン状態一括チェック開始")
    logger.info("==================================================")

    sessions_root = os.path.join(get_app_dir(), "browser_sessions")
    os.makedirs(sessions_root, exist_ok=True)
    accounts = load_accounts(include_paused=True)

    targets = []
    seen = set()
    for a in accounts:
        cid = a["clean_id"]
        if cid not in seen:
            seen.add(cid)
            targets.append((cid, a.get("auth_token", "")))
    for entry in sorted(os.listdir(sessions_root)):
        full = os.path.join(sessions_root, entry)
        if os.path.isdir(full):
            cid = entry.lstrip("_").lower()
            if cid not in seen:
                seen.add(cid)
                targets.append((cid, ""))

    results = {}
    status_path = os.path.join(get_app_dir(), "session_status.json")
    if os.path.exists(status_path):
        try:
            with open(status_path, "r", encoding="utf-8") as f:
                results = json.load(f)
        except Exception:
            results = {}

    fix_playwright_windows_driver()
    with sync_playwright() as p:
        for cid, tok in targets:
            if is_stop_requested():
                break
            sdir = resolve_session_dir(cid)
            status_label = "不明"
            detail = ""
            context = None
            try:
                kwargs = {
                    "user_data_dir": sdir,
                    "headless": True,
                    "viewport": {"width": 1280, "height": 800},
                    "args": ["--disable-blink-features=AutomationControlled"],
                }
                try:
                    context = p.chromium.launch_persistent_context(**kwargs, channel="chrome")
                except Exception:
                    context = p.chromium.launch_persistent_context(**kwargs)

                page = context.pages[0] if len(context.pages) > 0 else context.new_page()
                apply_media_blocking(page)
                if tok:
                    context.add_cookies([{"name": "auth_token", "value": tok, "domain": ".x.com", "path": "/"}])

                page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=25000)
                time.sleep(2.5)
                cur_url = (page.url or "").lower()
                susp = check_suspension_or_lock(page)
                if susp:
                    status_label = "凍結/制限"
                    detail = susp
                elif "login" in cur_url or "flow/login" in cur_url:
                    status_label = "要ログイン"
                    detail = "ログイン画面へリダイレクト"
                else:
                    status_label = "ログイン済み"
                    detail = "OK"
            except Exception as ce:
                err_s = str(ce)
                if "in use" in err_s.lower() or "lock" in err_s.lower() or "singleton" in err_s.lower():
                    status_label = "使用中(ロック)"
                else:
                    status_label = "確認エラー"
                detail = err_s[:80]
            finally:
                if context:
                    try:
                        context.close()
                    except Exception:
                        pass

            checked_at = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            results[cid] = {"status": status_label, "detail": detail, "checked_at": checked_at}
            logger.info(f"🔎 [{cid}] -> {status_label} ({detail})")

    try:
        with open(status_path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    logger.info("✅ 全プロファイルチェック完了！")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="X Browser Auto Tweet Bot")
    parser.add_argument("--test-account", type=str, default="", help="指定アカウントのみ1件テスト投稿")
    parser.add_argument("--test-theme", type=str, default="", help="テスト投稿時の対象テーマシート名")
    parser.add_argument("--check-sessions", action="store_true", help="全プロファイルのログイン状態をチェック")
    args, _unknown = parser.parse_known_args()

    try:
        if args.check_sessions:
            run_session_check()
        elif args.test_account:
            run_single_test(args.test_account, args.test_theme)
        else:
            safe_run()
    except Exception as e:
        logger.critical(f"重大なエラーが発生しました: {e}")
        logger.critical(traceback.format_exc())
        if not args.check_sessions and not args.test_account:
            print("\n" + "=" * 60)
            print("エラーが発生したため停止しました。")
            print("キーを押すと画面を閉じます...")
            print("=" * 60)
            input()
    finally:
        close_account_log()


if __name__ == "__main__":
    main()
