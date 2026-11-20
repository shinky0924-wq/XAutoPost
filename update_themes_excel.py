#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excelファイル テーマ別レイアウト移行＆見やすいカラーリング＋削除耐性書式設定ツール
--------------------------------------------------------------------
1. 不要になった「日常ツイート」「設定」シートを削除（旧「設定」の値はテーマ設定へ自動引き継ぎ）
2. 各テーマシートを以下の統一フォーマットへ整理・アップグレード:
   [1行目] 実行時間 | 曜日 | ポスト間隔 | 投稿モード | 有効/無効 | メモ
   [2行目] 0:00-0:00 | 毎日 | 120~360分 | 1: ランダム | 有効 | 投稿モード: 0=上から順 ...
   [3行目] （区切り空行）
   [4行目] No | ツイート文章 | 投稿時間 | 画像パス | ステータス | 投稿日時
   [5行目〜200行目] 投稿データ＆事前書式設定済みテンプレート行
3. セル値をDeleteキー等で削除・上書きしても罫線や背景色・ステータス色が消えないよう、
   200行目まで事前書式設定＋Excel条件付き書式（Conditional Formatting）を付与します。
"""

import os
import glob
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule


SETTINGS_HEADERS = ["実行時間", "曜日", "ポスト間隔", "投稿モード", "有効/無効", "メモ"]
DEFAULT_MEMO = (
    "投稿モード: 0=上から順 / 1=ランダム (数字だけでもOK) ｜ "
    "下のシートを追加・複製するとテーマ別に管理できます"
)
TWEET_HEADERS = ["No", "ツイート文章", "投稿時間", "画像パス", "ステータス", "投稿日時"]

TEMPLATE_MAX_ROW = 200


def _norm(val) -> str:
    if val is None:
        return ""
    return unicodedata.normalize("NFKC", str(val)).strip()


def _extract_legacy_settings(wb) -> dict:
    """旧「設定」シートがあれば値（実行時間・間隔・モード）を救出する"""
    res = {
        "time_range": "0:00-0:00",
        "weekdays": "毎日",
        "interval": "120~360分",
        "mode": "1: ランダム",
        "active": "有効",
    }
    if "設定" not in wb.sheetnames:
        return res

    ws = wb["設定"]
    try:
        c_min = ws.cell(3, 1).value
        c_max = None
        for col in range(2, ws.max_column + 1):
            if str(ws.cell(3, col).value or "").strip() in ("~", "〜", "～"):
                c_max = ws.cell(3, col + 1).value
                break
        if c_max is None:
            c_max = ws.cell(3, 4).value

        if c_min is not None:
            try:
                v_min = max(1, int(float(_norm(c_min))))
                v_max = max(v_min, int(float(_norm(c_max)))) if c_max is not None else v_min
                res["interval"] = f"{v_min}~{v_max}分"
            except Exception:
                pass

        for r in range(1, ws.max_row + 1):
            val = _norm(ws.cell(r, 1).value)
            if "時:分" in val or "ツイート実行時間" in val:
                target_r = r + 1 if "時:分" in val else r + 2
                s_val = _norm(ws.cell(target_r, 1).value)
                e_val = _norm(ws.cell(target_r, 3).value)
                if s_val and e_val:
                    res["time_range"] = f"{s_val}-{e_val}"
                break

        for r in range(4, ws.max_row + 1):
            val_text = _norm(ws.cell(r, 1).value)
            if "モード" in val_text and r + 2 <= ws.max_row:
                m_val = _norm(ws.cell(r + 2, 1).value)
                if m_val in ("0", "上から順"):
                    res["mode"] = "0: 上から順"
                elif m_val in ("1", "ランダム"):
                    res["mode"] = "1: ランダム"
                break
    except Exception:
        pass
    return res


def _is_theme_settings_row(ws) -> bool:
    a1 = _norm(ws.cell(1, 1).value)
    b1 = _norm(ws.cell(1, 2).value)
    c1 = _norm(ws.cell(1, 3).value)
    return "実行時間" in a1 and ("曜日" in b1 or "間隔" in c1)


def _find_tweet_header_row(ws) -> int | None:
    max_scan = min(15, max(ws.max_row, 1))
    for r in range(1, max_scan + 1):
        row_vals = [_norm(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 12) + 1)]
        for v in row_vals:
            if v in ("ツイート文章", "ポスト文章", "ツイート本文", "本文", "text"):
                return r
    return None


def _read_existing_sheet_data(ws, fallback_settings: dict) -> tuple[dict, list[dict]]:
    settings = dict(fallback_settings)

    has_theme_banner = _is_theme_settings_row(ws)
    if has_theme_banner:
        t_range = _norm(ws.cell(2, 1).value)
        wk = _norm(ws.cell(2, 2).value)
        iv = _norm(ws.cell(2, 3).value)
        md = _norm(ws.cell(2, 4).value)
        act = _norm(ws.cell(2, 5).value)
        memo = str(ws.cell(2, 6).value or "").strip()

        if t_range:
            settings["time_range"] = t_range
        if wk:
            settings["weekdays"] = wk
        if iv:
            settings["interval"] = iv
        if md:
            if md.startswith("0") or ("上" in md and "ランダム" not in md):
                settings["mode"] = "0: 上から順"
            elif md.startswith("1") or "ランダム" in md:
                settings["mode"] = "1: ランダム"
            else:
                settings["mode"] = md
        if act:
            settings["active"] = "無効" if act in ("無効", "停止", "OFF", "off", "0", "false", "FALSE") else "有効"
        if memo:
            settings["memo"] = memo

    header_row = _find_tweet_header_row(ws)
    tweets: list[dict] = []

    if header_row is not None:
        col_map = {"no": None, "text": None, "post_time": None, "image": None, "status": None, "posted_at": None}
        for c in range(1, ws.max_column + 1):
            h = _norm(ws.cell(header_row, c).value).lower()
            if h in ("no", "no.", "id", "番号"):
                col_map["no"] = c
            elif h in ("ツイート文章", "ポスト文章", "ツイート本文", "ポスト本文", "本文", "text"):
                col_map["text"] = c
            elif h in ("投稿時間", "投稿時間帯", "時間指定", "post_time", "time"):
                col_map["post_time"] = c
            elif h in ("画像パス", "画像", "画像フォルダ", "image_path", "image"):
                col_map["image"] = c
            elif h in ("ステータス", "status"):
                col_map["status"] = c
            elif h in ("投稿日時", "posted_at"):
                col_map["posted_at"] = c

        if col_map["text"] is None:
            col_map["text"] = 2

        for r in range(header_row + 1, ws.max_row + 1):
            raw_text = ws.cell(r, col_map["text"]).value
            text_str = str(raw_text).strip() if raw_text is not None else ""
            raw_img = ws.cell(r, col_map["image"]).value if col_map["image"] else ""
            img_str = str(raw_img).strip() if raw_img is not None else ""
            raw_pt = ws.cell(r, col_map["post_time"]).value if col_map["post_time"] else ""
            pt_str = str(raw_pt).strip() if raw_pt is not None else ""
            raw_st = ws.cell(r, col_map["status"]).value if col_map["status"] else ""
            st_str = str(raw_st).strip() if raw_st is not None else ""
            raw_dt = ws.cell(r, col_map["posted_at"]).value if col_map["posted_at"] else ""
            dt_str = str(raw_dt).strip() if raw_dt is not None else ""
            raw_no = ws.cell(r, col_map["no"]).value if col_map["no"] else None

            if not text_str and not img_str and not pt_str:
                continue

            tweets.append({
                "no": raw_no,
                "text": text_str,
                "post_time": pt_str,
                "image": img_str,
                "status": st_str,
                "posted_at": dt_str,
            })

    return settings, tweets


def _apply_theme_sheet_design(ws, settings: dict, tweets: list[dict]):
    """テーマシートにカラーデザイン＆削除耐性書式（200行事前書式＋条件付き書式）を適用"""
    # 既存の条件付き書式と結合セルをリセットしてクリーン適用
    ws.conditional_formatting._cf_rules.clear()
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    # 全セルいったんクリア（既存最大行まで）
    clear_until = max(ws.max_row, TEMPLATE_MAX_ROW)
    for r in range(1, clear_until + 1):
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    # ===== スタイル定義 =====
    border_thin_slate = Side(border_style="thin", color="CBD5E1")
    border_med_navy = Side(border_style="medium", color="1E3A8A")
    border_thin_navy = Side(border_style="thin", color="93C5FD")

    grid_border = Border(
        left=border_thin_slate, right=border_thin_slate,
        top=border_thin_slate, bottom=border_thin_slate,
    )
    settings_header_border = Border(
        left=Side(border_style="thin", color="3B82F6"),
        right=Side(border_style="thin", color="3B82F6"),
        top=border_med_navy,
        bottom=border_thin_navy,
    )
    settings_val_border = Border(
        left=border_thin_navy, right=border_thin_navy,
        top=border_thin_navy, bottom=border_med_navy,
    )
    table_header_border = Border(
        left=Side(border_style="thin", color="475569"),
        right=Side(border_style="thin", color="475569"),
        top=border_med_navy,
        bottom=border_med_navy,
    )

    fill_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_setting_blue = PatternFill(start_color="E8F1FF", end_color="E8F1FF", fill_type="solid")
    fill_active_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_inactive_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_memo = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    fill_table_hdr = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
    fill_time_hdr = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")

    fill_row_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_row_even = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_time_empty_odd = PatternFill(start_color="FFFDF5", end_color="FFFDF5", fill_type="solid")
    fill_time_empty_even = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
    fill_time_set = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    fill_status_posted = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_status_pending = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_status_failed = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_status_skipped = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

    font_hdr = Font(name="Yu Gothic UI", size=11, bold=True, color="FFFFFF")
    font_setting_val = Font(name="Yu Gothic UI", size=11, bold=True, color="1E3A8A")
    font_active_green = Font(name="Yu Gothic UI", size=11, bold=True, color="15803D")
    font_inactive_red = Font(name="Yu Gothic UI", size=11, bold=True, color="B91C1C")
    font_memo = Font(name="Yu Gothic UI", size=9, bold=False, color="475569")

    font_no = Font(name="Yu Gothic UI", size=10, bold=True, color="64748B")
    font_body = Font(name="Yu Gothic UI", size=10, bold=False, color="0F172A")
    font_time = Font(name="Yu Gothic UI", size=10, bold=True, color="92400E")
    font_img = Font(name="Yu Gothic UI", size=10, bold=False, color="334155")
    font_posted = Font(name="Yu Gothic UI", size=10, bold=True, color="15803D")
    font_pending = Font(name="Yu Gothic UI", size=10, bold=True, color="B45309")
    font_failed = Font(name="Yu Gothic UI", size=10, bold=True, color="B91C1C")
    font_skipped = Font(name="Yu Gothic UI", size=10, bold=False, color="475569")
    font_date = Font(name="Yu Gothic UI", size=10, bold=False, color="475569")

    # ===== Row 1: テーマ設定ヘッダー =====
    for c_idx, h_text in enumerate(SETTINGS_HEADERS, 1):
        cell = ws.cell(1, c_idx, h_text)
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = settings_header_border
    ws.row_dimensions[1].height = 28

    # ===== Row 2: テーマ設定値 =====
    active_val = settings.get("active", "有効")
    is_enabled = active_val not in ("無効", "停止", "OFF", "off", "0", "false", "FALSE")
    row2_values = [
        settings.get("time_range", "0:00-0:00"),
        settings.get("weekdays", "毎日"),
        settings.get("interval", "120~360分"),
        settings.get("mode", "1: ランダム"),
        "有効" if is_enabled else "無効",
        settings.get("memo", DEFAULT_MEMO),
    ]
    for c_idx, val in enumerate(row2_values, 1):
        cell = ws.cell(2, c_idx, val)
        cell.border = settings_val_border
        if c_idx <= 4:
            cell.fill = fill_setting_blue
            cell.font = font_setting_val
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        elif c_idx == 5:
            cell.fill = fill_active_green if is_enabled else fill_inactive_red
            cell.font = font_active_green if is_enabled else font_inactive_red
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.fill = fill_memo
            cell.font = font_memo
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 42

    # ===== Row 3: 空行スペーサー =====
    ws.row_dimensions[3].height = 10

    # ===== Row 4: 投稿テーブルヘッダー =====
    for c_idx, h_text in enumerate(TWEET_HEADERS, 1):
        cell = ws.cell(4, c_idx, h_text)
        cell.font = font_hdr
        cell.fill = fill_time_hdr if c_idx == 3 else fill_table_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = table_header_border
    ws.row_dimensions[4].height = 26

    # ===== Row 5〜max(200, len+30): データ＆事前書式行 =====
    total_rows = max(TEMPLATE_MAX_ROW, len(tweets) + 30)
    for idx in range(total_rows - 4):
        r = 5 + idx
        is_even = (r % 2 == 0)
        base_fill = fill_row_even if is_even else fill_row_odd
        time_empty_fill = fill_time_empty_even if is_even else fill_time_empty_odd

        t_item = tweets[idx] if idx < len(tweets) else None
        no_val = (t_item["no"] if t_item and t_item["no"] not in (None, "") else idx + 1) if t_item else None
        text_val = t_item["text"] if t_item else None
        pt_val = t_item["post_time"] if (t_item and t_item["post_time"]) else None
        img_val = t_item["image"] if (t_item and t_item["image"]) else None
        st_val = t_item["status"] if (t_item and t_item["status"]) else ("pending" if t_item else None)
        dt_val = t_item["posted_at"] if (t_item and t_item["posted_at"]) else None

        # A: No
        c_no = ws.cell(r, 1, no_val)
        c_no.font = font_no
        c_no.fill = base_fill
        c_no.alignment = Alignment(horizontal="center", vertical="center")
        c_no.border = grid_border

        # B: ツイート文章
        c_txt = ws.cell(r, 2, text_val)
        c_txt.font = font_body
        c_txt.fill = base_fill
        c_txt.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_txt.border = grid_border

        # C: 投稿時間
        c_pt = ws.cell(r, 3, pt_val)
        c_pt.font = font_time
        c_pt.fill = fill_time_set if pt_val else time_empty_fill
        c_pt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_pt.border = grid_border

        # D: 画像パス
        c_img = ws.cell(r, 4, img_val)
        c_img.font = font_img
        c_img.fill = base_fill
        c_img.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_img.border = grid_border

        # E: ステータス
        c_st = ws.cell(r, 5, st_val)
        c_st.border = grid_border
        c_st.alignment = Alignment(horizontal="center", vertical="center")
        st_lower = str(st_val or "").strip().lower()
        if st_lower in ("posted", "投稿済み", "投稿済", "完了"):
            c_st.fill = fill_status_posted
            c_st.font = font_posted
        elif st_lower in ("pending", "未投稿"):
            c_st.fill = fill_status_pending
            c_st.font = font_pending
        elif st_lower in ("failed", "失敗", "エラー"):
            c_st.fill = fill_status_failed
            c_st.font = font_failed
        elif st_lower in ("skipped", "スキップ"):
            c_st.fill = fill_status_skipped
            c_st.font = font_skipped
        else:
            c_st.fill = base_fill
            c_st.font = font_pending

        # F: 投稿日時
        c_dt = ws.cell(r, 6, dt_val)
        c_dt.font = font_date
        c_dt.fill = base_fill
        c_dt.alignment = Alignment(horizontal="center", vertical="center")
        c_dt.border = grid_border

        if t_item and text_val and ("\n" in text_val or len(text_val) > 48):
            lines = max(text_val.count("\n") + 1, (len(text_val) // 36) + 1)
            ws.row_dimensions[r].height = min(140, max(24, lines * 16 + 6))
        else:
            ws.row_dimensions[r].height = 24

    # ===== 列幅設定 =====
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 54
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 38

    # ヘッダー固定＆グリッド線表示
    ws.freeze_panes = "A5"
    ws.views.sheetView[0].showGridLines = True

    # ===== 条件付き書式（セルをDelete・ペーストしても色が消えない保護）=====
    last_r = total_rows

    # 1. 有効/無効 (E2)
    ws.conditional_formatting.add(
        "E2",
        CellIsRule(operator="equal", formula=['"無効"'], fill=fill_inactive_red, font=font_inactive_red),
    )
    ws.conditional_formatting.add(
        "E2",
        CellIsRule(operator="equal", formula=['"有効"'], fill=fill_active_green, font=font_active_green),
    )

    # 2. ステータス列 (E5:E200)
    st_rng = f"E5:E{last_r}"
    for word in ['"posted"', '"投稿済"', '"投稿済み"', '"完了"']:
        ws.conditional_formatting.add(
            st_rng,
            CellIsRule(operator="equal", formula=[word], fill=fill_status_posted, font=font_posted),
        )
    for word in ['"pending"', '"未投稿"']:
        ws.conditional_formatting.add(
            st_rng,
            CellIsRule(operator="equal", formula=[word], fill=fill_status_pending, font=font_pending),
        )
    for word in ['"failed"', '"失敗"', '"エラー"']:
        ws.conditional_formatting.add(
            st_rng,
            CellIsRule(operator="equal", formula=[word], fill=fill_status_failed, font=font_failed),
        )
    for word in ['"skipped"', '"スキップ"']:
        ws.conditional_formatting.add(
            st_rng,
            CellIsRule(operator="equal", formula=[word], fill=fill_status_skipped, font=font_skipped),
        )
    # ステータス空欄時もゼブラ背景を維持
    ws.conditional_formatting.add(
        st_rng,
        FormulaRule(formula=['AND(ISBLANK(E5),MOD(ROW(),2)=0)'], fill=fill_row_even),
    )
    ws.conditional_formatting.add(
        st_rng,
        FormulaRule(formula=['AND(ISBLANK(E5),MOD(ROW(),2)=1)'], fill=fill_row_odd),
    )

    # 3. 投稿時間列 (C5:C200): 値が入っていれば強調色、空欄なら薄いクリームゼブラ
    pt_rng = f"C5:C{last_r}"
    ws.conditional_formatting.add(
        pt_rng,
        FormulaRule(formula=['LEN(TRIM(C5))>0'], fill=fill_time_set, font=font_time),
    )
    ws.conditional_formatting.add(
        pt_rng,
        FormulaRule(formula=['AND(LEN(TRIM(C5))=0,MOD(ROW(),2)=0)'], fill=fill_time_empty_even),
    )
    ws.conditional_formatting.add(
        pt_rng,
        FormulaRule(formula=['AND(LEN(TRIM(C5))=0,MOD(ROW(),2)=1)'], fill=fill_time_empty_odd),
    )

    # 4. 本体・画像・日時列 (A, B, D, F) のゼブラ保護
    for col_rng, top_cell in [(f"A5:B{last_r}", "A5"), (f"D5:D{last_r}", "D5"), (f"F5:F{last_r}", "F5")]:
        ws.conditional_formatting.add(
            col_rng,
            FormulaRule(formula=["MOD(ROW(),2)=0"], fill=fill_row_even),
        )
        ws.conditional_formatting.add(
            col_rng,
            FormulaRule(formula=["MOD(ROW(),2)=1"], fill=fill_row_odd),
        )


def migrate_workbook_to_themes(file_path: str) -> bool:
    try:
        wb = openpyxl.load_workbook(file_path)
        legacy_settings = _extract_legacy_settings(wb)

        # 旧シート「日常ツイート」「設定」を削除
        removed_sheets = []
        for legacy_name in ("日常ツイート", "設定"):
            if legacy_name in wb.sheetnames:
                wb.remove(wb[legacy_name])
                removed_sheets.append(legacy_name)

        # もし全てのシートが消えた場合はデフォルト「テーマ1」を作成
        if not wb.sheetnames:
            wb.create_sheet(title="テーマ1")

        # 残ったすべてのシート（テーマ）を新レイアウト・カラー書式に統一
        processed_themes = []
        for s_name in list(wb.sheetnames):
            ws = wb[s_name]
            settings, tweets = _read_existing_sheet_data(ws, legacy_settings)
            _apply_theme_sheet_design(ws, settings, tweets)
            processed_themes.append(f"{s_name}({len(tweets)}件)")

        wb.save(file_path)
        rem_info = f" / 削除シート: {', '.join(removed_sheets)}" if removed_sheets else ""
        print(f"✅ テーマ最適化完了: {os.path.basename(file_path)} [{', '.join(processed_themes)}{rem_info}]")
        return True
    except Exception as e:
        print(f"⚠️ Excel更新エラー ({os.path.basename(file_path)}): {e}")
        return False


def main():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    onedrive_desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
    onedrive_jp = os.path.join(os.path.expanduser("~"), "OneDrive", "デスクトップ")

    candidate_dirs = [
        ".",
        os.path.join(desktop, "X_BOT_配布用(Python不要)"),
        os.path.join(onedrive_desktop, "X_BOT_配布用(Python不要)"),
        os.path.join(onedrive_jp, "X_BOT_配布用(Python不要)"),
        r"C:\Users\PC_User\Desktop\X_BOT_配布用(Python不要)",
        r"C:\Users\PC_User\Desktop\X_BOT",
    ]

    visited = set()
    for d in candidate_dirs:
        if not d or not os.path.exists(d):
            continue
        abs_d = os.path.abspath(d)
        if abs_d in visited:
            continue
        visited.add(abs_d)

        for xlsx in glob.glob(os.path.join(abs_d, "*.xlsx")):
            fname = os.path.basename(xlsx)
            if "アカウント情報" in fname or "dm_targets" in fname or fname.startswith("~$"):
                continue
            migrate_workbook_to_themes(xlsx)


if __name__ == "__main__":
    main()
