#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================
X (Twitter) ブラウザ自動操縦・自動ツイートボット (完全統合・最新版)
====================================================================
【統合機能一覧】
  1. 【提案1】1窓・順番交代ローテーション (常に開くブラウザは1つのみ)
  2. 【提案2】画像・動画通信ブロック ＋ アップロード通信保護 (メモリ約60%削減)
  3. 【日常ツイート独立エンジン】朝枠・昼枠・夕方枠・夜枠の自動判定＆30日クールダウン
  4. 【スピンタックス＆動的置換】{A|B|C} ＆ {月日}{曜日}{時間}{日付}
  5. 【案3: 画像クールダウン＆アカウント間重複防止】
     - 同一アカウント：30日間再選出禁止 (image_history.json)
     - 他アカウント：7日間共有禁止（同一IP重複対策）
  6. 【凍結アカウント即時永久除外】
     - 凍結・制限検知時に即座にブラックリスト追加＆Excel除外、再起動不要で継続
  7. 【chrome_launcher.exe 連携】
     - 手動ログインで作成したプロファイル (browser_sessions/) を自動引き継ぎ
  8. 【括弧付きフォルダ対応】
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
import openpyxl
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] [XBrowserBot] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("browser_auto_tweet.log", encoding="utf-8")
    ]
)
logger = logging.getLogger("XBrowserBot")


def get_app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


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


# 朝枠・昼枠・夕方枠・夜枠のフォルダからクールダウンを判定して選定
def resolve_random_image_from_folder(raw_target, slot_name="", handle=""):
    target_path = str(raw_target or "").strip().replace('"', '').replace("'", "")
    if target_path in ["(なし)", "なし", "none", "None"]:
        return None

    app_dir = get_app_dir()
    search_dirs = [
        target_path,
        os.path.join("images", target_path),
        os.path.join("images", slot_name) if slot_name else "",
        os.path.join(app_dir, target_path) if target_path else "",
        os.path.join(app_dir, "images", target_path) if target_path else "",
        os.path.join(app_dir, "images", slot_name) if slot_name else "",
    ]

    if target_path and os.path.isfile(target_path):
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

    # クールダウン（自垢30日、他垢7日）をクリアした画像のみを候補にする
    available_files = []
    for fpath in candidates:
        fname = os.path.basename(fpath)
        ok, _ = is_image_available(fname, handle, history)
        if ok:
            available_files.append(fpath)

    if available_files:
        chosen = random.choice(available_files)
        logger.info(f"📸 画像フォルダ [{os.path.basename(valid_folder)}] から選出（クールダウン合格: {os.path.basename(chosen)} / 候補数: {len(available_files)}/{len(candidates)}）")
        return os.path.abspath(chosen)

    # 全画像がクールダウン中の場合、最も昔に使われた画像を安全にフォールバック選出
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


def parse_time_str(val):
    if not val:
        return None
    s = str(val).strip()
    try:
        parts = s.split(":")
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        return h * 60 + m
    except Exception:
        return None


def is_current_time_in_range(time_range_str):
    if not time_range_str:
        return False
    try:
        s_str, e_str = str(time_range_str).split("-")
        s_min = parse_time_str(s_str.strip())
        e_min = parse_time_str(e_str.strip())
        if s_min is None or e_min is None:
            return False

        now = datetime.now()
        cur_min = now.hour * 60 + now.minute
        if s_min < e_min:
            return s_min <= cur_min < e_min
        else:
            return cur_min >= s_min or cur_min < e_min
    except Exception:
        return False


# 日常ツイートの選出（30日クールダウン＆画像フォルダ選定）
def get_eligible_daily_tweet(excel_path, handle):
    if not excel_path or not os.path.exists(excel_path):
        return None
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "日常ツイート" not in wb.sheetnames:
            return None
        ws = wb["日常ツイート"]

        now = datetime.now()
        today_str = now.strftime("%Y/%m/%d")
        cooldown_days = 30
        candidates = []

        for r in range(2, ws.max_row + 1):
            slot_name = str(ws.cell(r, 1).value or "").strip()
            time_range = str(ws.cell(r, 2).value or "").strip()
            raw_text = str(ws.cell(r, 3).value or "").strip()
            status = str(ws.cell(r, 4).value or "").strip()
            last_date_val = str(ws.cell(r, 5).value or "").strip()
            raw_img_folder = str(ws.cell(r, 7).value or "").strip()

            if status not in ["有効", "active", "1", "ON", "on"]:
                continue
            if not is_current_time_in_range(time_range):
                continue
            if last_date_val.startswith(today_str):
                continue

            is_cooldown_ok = True
            if last_date_val:
                try:
                    last_dt = datetime.strptime(last_date_val[:10], "%Y/%m/%d")
                    if (now - last_dt).days < cooldown_days:
                        is_cooldown_ok = False
                except Exception:
                    pass
            if not is_cooldown_ok:
                continue

            candidates.append({
                "row": r,
                "slot_name": slot_name,
                "raw_text": raw_text,
                "raw_img": raw_img_folder
            })

        if candidates:
            chosen = random.choice(candidates)
            final_text = expand_spintax(chosen["raw_text"])
            final_text = replace_dynamic_tags(final_text)
            chosen["final_text"] = final_text
            chosen["image"] = resolve_random_image_from_folder(chosen["raw_img"], chosen["slot_name"], handle)
            return chosen
    except Exception as e:
        logger.error(f"日常ツイート取得エラー: {e}")
    return None


def update_daily_tweet_status(excel_path, row_idx):
    now = datetime.now()
    for attempt in range(3):
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb["日常ツイート"]
            ws.cell(row_idx, 5).value = now.strftime("%Y/%m/%d")
            ws.cell(row_idx, 6).value = now.strftime("%Y/%m/%d %H:%M:%S")
            wb.save(excel_path)
            logger.info(f"日常ツイート更新成功: 行{row_idx} (最終投稿日更新)")
            return
        except PermissionError:
            time.sleep(2)
        except Exception:
            break


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
    if os.path.isdir(clean):
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


def load_accounts():
    accounts_file = "アカウント情報.xlsx"
    accounts = []
    if not os.path.exists(accounts_file):
        return accounts
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


def get_settings(excel_path):
    val_min = 60
    val_max = 120
    settings = {
        "min_interval": 60, "max_interval": 120, "mode": 0,
        "start_time_min": None, "end_time_min": None
    }
    if not excel_path or not os.path.exists(excel_path):
        return settings
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "設定" in wb.sheetnames:
            ws = wb["設定"]
            c_min = ws.cell(3, 1).value
            c_max = None
            for col in range(2, ws.max_column + 1):
                if str(ws.cell(3, col).value or "").strip() == "~":
                    c_max = ws.cell(3, col + 1).value
                    break
            if c_max is None:
                c_max = ws.cell(3, 4).value

            try:
                if c_min is not None:
                    val_min = int(float(str(c_min).strip()) * 60)
            except Exception:
                val_min = 60
            try:
                if c_max is not None:
                    val_max = int(float(str(c_max).strip()) * 60)
            except Exception:
                val_max = val_min

            final_min = max(1, min(val_min, val_max))
            final_max = max(val_min, val_max)
            settings["min_interval"] = final_min
            settings["max_interval"] = final_max

            for r in range(1, ws.max_row + 1):
                val = str(ws.cell(r, 1).value or "")
                if "時:分" in val or "ツイート実行時間" in val:
                    target_r = r + 1 if "時:分" in val else r + 2
                    s_min = parse_time_str(ws.cell(target_r, 1).value)
                    e_min = parse_time_str(ws.cell(target_r, 3).value)
                    if s_min is not None and e_min is not None and s_min != e_min:
                        settings["start_time_min"] = s_min
                        settings["end_time_min"] = e_min
                    break

            for r in range(4, ws.max_row + 1):
                val_text = str(ws.cell(r, 1).value or "")
                if "モード" in val_text:
                    mode_val = ws.cell(r + 2, 1).value
                    if mode_val is not None and str(mode_val).strip() in ["1", "ランダム"]:
                        settings["mode"] = 1
                    break
    except Exception as e:
        logger.warning(f"設定取得エラー: {e}")
    return settings


def is_within_operating_hours(settings):
    s_min = settings.get("start_time_min")
    e_min = settings.get("end_time_min")
    if s_min is None or e_min is None:
        return True
    now = datetime.now()
    cur_min = now.hour * 60 + now.minute
    if s_min < e_min:
        return s_min <= cur_min < e_min
    else:
        return cur_min >= s_min or cur_min < e_min


def reset_all_tweets_status(excel_path):
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb["ツイート文章"] if "ツイート文章" in wb.sheetnames else wb.active
        headers = [str(cell.value or "").strip() for cell in ws[1]]
        status_col = 4
        for idx, h in enumerate(headers):
            if h in ["ステータス", "status"]:
                status_col = idx + 1
                break
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 2).value:
                ws.cell(r, status_col).value = ""
        wb.save(excel_path)
        logger.info(f"[{os.path.basename(excel_path)}] 全ツイート完了のためステータス自動リセット。")
    except Exception:
        pass


def load_tweets(excel_path, handle=""):
    tweets = []
    if not excel_path or not os.path.exists(excel_path):
        return tweets
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb["ツイート文章"] if "ツイート文章" in wb.sheetnames else wb.active
        headers = [str(cell.value or "").strip() for cell in ws[1]]
        text_col, img_col, status_col, date_col = 2, 3, 4, 5
        for idx, h in enumerate(headers):
            if h in ["ツイート文章", "text", "本文"]: text_col = idx + 1
            elif h in ["画像パス", "画像"]: img_col = idx + 1
            elif h in ["ステータス", "status"]: status_col = idx + 1
            elif h in ["投稿日時"]: date_col = idx + 1

        for r in range(2, ws.max_row + 1):
            text = str(ws.cell(r, text_col).value or "").strip()
            if not text:
                continue
            raw_img = str(ws.cell(r, img_col).value or "").strip()
            resolved_img = resolve_image_path(raw_img, handle=handle)
            status = str(ws.cell(r, status_col).value or "").strip().lower()
            if status not in ["posted", "投稿済み", "完了", "skipped"]:
                tweets.append({
                    "row": r, "text": text, "image": resolved_img,
                    "status_col": status_col, "date_col": date_col
                })
    except Exception as e:
        logger.error(f"ツイート読み込みエラー: {e}")
    return tweets


def update_excel_status(excel_path, row_idx, status_col, date_col, status_text):
    for attempt in range(3):
        try:
            wb = openpyxl.load_workbook(excel_path)
            ws = wb["ツイート文章"] if "ツイート文章" in wb.sheetnames else wb.active
            ws.cell(row_idx, status_col).value = status_text
            ws.cell(row_idx, date_col).value = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            wb.save(excel_path)
            logger.info(f"Excel更新完了: 行{row_idx} -> {status_text}")
            return
        except PermissionError:
            time.sleep(2)
        except Exception:
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
            if dup in content: return "duplicate"
        for err in error_texts:
            if err in content: return "general_error"
    except Exception:
        pass
    return None


def execute_single_post_attempt(page, text_to_send, image_path):
    try:
        page.goto("https://x.com/compose/post", wait_until="domcontentloaded", timeout=30000)
        time.sleep(random.uniform(2.0, 3.0))

        susp_msg = check_suspension_or_lock(page)
        if susp_msg:
            return f"suspended:{susp_msg}"

        human_mouse_move(page)
        editor = page.locator('div[role="textbox"]').first
        try:
            editor.wait_for(state="visible", timeout=8000)
        except Exception:
            susp_msg = check_suspension_or_lock(page)
            if susp_msg:
                return f"suspended:{susp_msg}"
            return "general_error"

        box = editor.bounding_box()
        if box:
            page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2, steps=4)
            time.sleep(0.1)
            page.mouse.down()
            time.sleep(0.05)
            page.mouse.up()
        else:
            editor.click()

        time.sleep(random.uniform(0.3, 0.6))
        human_typing(page, text_to_send)
        time.sleep(random.uniform(0.8, 1.5))

        if image_path and os.path.exists(image_path):
            try:
                logger.info(f"📸 画像ファイルを添付中: {os.path.basename(image_path)}")
                file_input = page.locator("input[type='file'][data-testid='fileInput']").first
                file_input.set_input_files(image_path)
                time.sleep(random.uniform(3.5, 5.0))
            except Exception as img_err:
                logger.warning(f"画像添付エラー: {img_err}")

        time.sleep(random.uniform(0.5, 1.0))

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
            return f"suspended:{susp_msg}"

        err_type = check_for_errors_or_duplicates(page)
        if err_type:
            return err_type
        return "success"
    except Exception as e:
        logger.error(f"投稿実行エラー: {e}")
        return "general_error"


def post_tweet_with_smart_retry(page, original_text, image_path):
    res1 = execute_single_post_attempt(page, original_text, image_path)
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

    res2 = execute_single_post_attempt(page, spaced_text, image_path)
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
    """
    括弧・日本語を含まないパス（%LOCALAPPDATA%\\x_bot_pw_driver 等）に
    NTFSディレクトリジャンクションを作成し、ドライバディレクトリをクリーンに参照可能にする。
    """
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
        # パスに日本語等が含まれていても問題ない純ASCIIパスか優先確認
        junction_path = os.path.join(parent, "x_bot_pw_driver")
        try:
            # 既存のジャンクションまたはリンク先が古い場合は作り直す
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
    """ジャンクションが使えない場合のフォールバック用 .cmd ラッパー"""
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
    """
    Windowsでフォルダ名に括弧「( )」やスペースが含まれる場合（例: X_BOT_配布用(Python不要)）、
    Playwrightドライバ起動時に cmd.exe がパスを途中で切断する問題を完全に回避する。
    """
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

        # 実ドライバディレクトリを特定し、括弧等がある場合はクリーンなジャンクションを作成
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
                fixed = tuple(_rewrite_path(x) for x in res)
                return fixed
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

        # 最終防衛ライン: asyncio.create_subprocess_exec でも括弧付き .cmd を回避
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


def safe_run():
    logger.info("==================================================")
    logger.info("★ X 自動操縦ボット起動 (画像クールダウン＆日常ツイート完全統合版)")
    logger.info("★ 案3: 画像クールダウン（同一垢30日 / 他垢7日共有ロック）")
    logger.info("★ 日常ツイート: 朝枠・昼枠・夕方枠・夜枠のフォルダから画像自動選出")
    logger.info("★ 1窓交代ローテーション ＆ 画像・動画通信ブロック(メモリ60%カット)")
    logger.info("★ 凍結対策: 凍結検知時にメモリから即座に永久排除 ＆ 再起動不要で完全自動継続！")
    logger.info(f"★ 登録済みブラックリスト除外アカウント数: {len(BLACKLISTED_ACCOUNTS)} 件")
    logger.info("==================================================")

    next_post_times = {}
    fix_playwright_windows_driver()

    with sync_playwright() as p:
        while True:
            accounts = load_accounts()
            if not accounts:
                logger.warning("運用可能な健全アカウントが見つかりません。30秒後に再確認します。")
                time.sleep(30)
                continue

            now = datetime.now()
            worked_in_this_cycle = False

            for acc in accounts:
                handle, clean_id = acc["handle"], acc["clean_id"]
                if clean_id in BLACKLISTED_ACCOUNTS or not acc.get("excel_path"):
                    continue

                settings = get_settings(acc["excel_path"])
                excel_path = acc["excel_path"]

                # 1. まず日常ツイート枠があるかチェック
                daily_item = get_eligible_daily_tweet(excel_path, handle)
                is_daily, target_text, target_img = False, "", None

                if daily_item:
                    is_daily = True
                    target_text = daily_item["final_text"]
                    target_img = daily_item.get("image")
                    logger.info(f"✨ [{handle}] 日常ツイート枠「{daily_item['slot_name']}」を検知しました！")
                else:
                    if not is_within_operating_hours(settings):
                        continue
                    if handle in next_post_times and now < next_post_times[handle]:
                        continue
                    tweets = load_tweets(excel_path, handle=handle)
                    if not tweets:
                        reset_all_tweets_status(excel_path)
                        tweets = load_tweets(excel_path, handle=handle)
                    if not tweets:
                        continue
                    target_tweet = random.choice(tweets) if settings["mode"] == 1 else tweets[0]
                    target_text, target_img = target_tweet["text"], target_tweet["image"]

                mode_label = "【日常ツイート】" if is_daily else "【通常ツイート】"
                img_label = f" (画像: {os.path.basename(target_img)})" if target_img else " (画像なし)"
                logger.info("\n--------------------------------------------------")
                logger.info(f"🔄 [{handle}] {mode_label}{img_label}")
                logger.info(f"📝 {target_text[:35]}...")
                logger.info("--------------------------------------------------")

                session_dir = resolve_session_dir(handle)
                logger.info(f"📂 プロファイル参照先: {session_dir}")

                context, is_suspended, suspension_reason = None, False, ""
                try:
                    context = launch_browser_safely(p, session_dir)
                    page = context.pages[0] if len(context.pages) > 0 else context.new_page()
                    apply_media_blocking(page)

                    if acc.get("auth_token"):
                        context.add_cookies([{"name": "auth_token", "value": acc["auth_token"], "domain": ".x.com", "path": "/"}])

                    page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=30000)
                    time.sleep(random.uniform(2.0, 3.0))

                    susp_check = check_suspension_or_lock(page)
                    if susp_check:
                        is_suspended, suspension_reason = True, susp_check
                    else:
                        res = post_tweet_with_smart_retry(page, target_text, target_img)
                        if res.startswith("suspended:"):
                            is_suspended, suspension_reason = True, res.replace("suspended:", "")
                        elif res == "success":
                            # 投稿成功時に画像履歴を永続記録（案3 クールダウン30日・他垢7日）
                            if target_img:
                                record_image_usage(target_img, handle)

                            if is_daily:
                                update_daily_tweet_status(excel_path, daily_item["row"])
                            else:
                                update_excel_status(excel_path, target_tweet["row"], target_tweet["status_col"], target_tweet["date_col"], "posted")
                        else:
                            if not is_daily:
                                update_excel_status(excel_path, target_tweet["row"], target_tweet["status_col"], target_tweet["date_col"], "skipped")
                except Exception as loop_e:
                    logger.error(f"[{handle}] エラー: {loop_e}")
                finally:
                    if context:
                        try:
                            context.close()
                            logger.info(f"🧹 [{handle}] のブラウザを閉じ、メモリを全解放しました。")
                        except Exception:
                            pass

                if is_suspended:
                    logger.critical(f"🚨 [{handle}] の凍結検知（原因: {suspension_reason}）")
                    blacklist_and_remove_account(handle, reason=suspension_reason)
                    if handle in next_post_times:
                        del next_post_times[handle]
                    logger.info("⚡ 再起動せず即時排除しました。残りのアカウントで継続します。")
                    time.sleep(2)
                    break

                t_min = min(settings["min_interval"], settings["max_interval"])
                t_max = max(settings["min_interval"], settings["max_interval"])
                wait_sec = random.randint(t_min, t_max) if t_min < t_max else t_min
                next_post_times[handle] = datetime.now() + timedelta(seconds=wait_sec)

                logger.info(f"⏳ [{handle}] 次回: {next_post_times[handle].strftime('%H:%M:%S')} ({wait_sec // 60}分 {wait_sec % 60}秒 後)")
                worked_in_this_cycle = True
                time.sleep(random.uniform(5.0, 10.0))

            if not worked_in_this_cycle:
                time.sleep(15)


def main():
    try:
        safe_run()
    except Exception as e:
        logger.critical(f"重大なエラーが発生しました: {e}")
        logger.critical(traceback.format_exc())
        print("\n" + "=" * 60)
        print("エラーが発生したため停止しました。")
        print("キーを押すと画面を閉じます...")
        print("=" * 60)
        input()


if __name__ == "__main__":
    main()
