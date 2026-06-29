#!/usr/bin/env python3
"""
個別相談トラッカー
文字起こしから顧客情報を整理し、Excelに出力する
"""

import json
import os
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
CONSULTATIONS_FILE = BASE_DIR / "consultations.json"
EXCEL_OUTPUT = BASE_DIR / "consultation_tracker.xlsx"


# -----------------------------------------------------------------
# データ構造
# -----------------------------------------------------------------

EMPTY_RECORD = {
    "id": None,
    "name": "",                    # 名前（ニックネーム可）
    "date_consulted": "",          # 相談日 (YYYY-MM-DD)
    "current_score": {
        "overall": "",             # 総合スコア
        "listening": "",
        "reading": "",
        "writing": "",
        "speaking": "",
    },
    "target_score": {
        "overall": "",
        "listening": "",
        "reading": "",
        "writing": "",
        "speaking": "",
    },
    "last_exam_date": "",          # 最後にIELTSを受験した日
    "next_exam_date": "",          # 次回受験予定日
    "exam_type": "",               # Academic / General
    "exam_purpose": "",            # 受験目的（留学・移住・就職など）
    "target_school_country": "",   # 目標の学校・国
    "deadline": "",                # スコア提出締め切り
    "study_situation": "",         # 現在の学習状況
    "weak_points": "",             # 弱点・課題
    "concerns": "",                # 不安・悩み
    "schedule_plan": "",           # 学習スケジュール・計画
    "notes": "",                   # その他メモ
}


# -----------------------------------------------------------------
# JSON 読み書き
# -----------------------------------------------------------------

def load_consultations() -> list:
    if not CONSULTATIONS_FILE.exists():
        return []
    with open(CONSULTATIONS_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_consultations(records: list) -> None:
    with open(CONSULTATIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


# -----------------------------------------------------------------
# 入力補助
# -----------------------------------------------------------------

def ask(prompt: str, default: str = "") -> str:
    val = input(f"  {prompt}" + (f" [{default}]" if default else "") + ": ").strip()
    return val if val else default


def ask_score(label: str) -> dict:
    print(f"\n  【{label}スコア】（不明な場合は空でEnter）")
    return {
        "overall":   ask("総合"),
        "listening": ask("Listening"),
        "reading":   ask("Reading"),
        "writing":   ask("Writing"),
        "speaking":  ask("Speaking"),
    }


# -----------------------------------------------------------------
# 新規入力
# -----------------------------------------------------------------

def input_new_record(records: list) -> dict:
    next_id = max((r["id"] for r in records), default=0) + 1
    rec = {**EMPTY_RECORD}
    rec["id"] = next_id

    print(f"\n  ── 新規相談者 ID:{next_id} ──")
    rec["name"]           = ask("名前（ニックネーム可）")
    rec["date_consulted"] = ask("相談日 (YYYY-MM-DD)", datetime.today().strftime("%Y-%m-%d"))
    rec["exam_type"]      = ask("受験タイプ (Academic / General)", "Academic")
    rec["exam_purpose"]   = ask("受験目的（例: UK大学院留学、カナダ移住）")
    rec["target_school_country"] = ask("目標の学校・国")
    rec["deadline"]       = ask("スコア提出締め切り (YYYY-MM-DD)")

    rec["current_score"]  = ask_score("現在の")
    rec["target_score"]   = ask_score("目標")

    print()
    rec["last_exam_date"] = ask("最後にIELTSを受験した日 (YYYY-MM-DD、未受験は空)")
    rec["next_exam_date"] = ask("次回受験予定日 (YYYY-MM-DD)")

    print()
    rec["study_situation"] = ask("現在の学習状況（例: 毎日30分、単語帳中心）")
    rec["weak_points"]     = ask("弱点・課題（例: Writingのタスク達成度）")
    rec["concerns"]        = ask("不安・悩み（相談者の言葉をそのまま）")
    rec["schedule_plan"]   = ask("学習スケジュール・計画（例: 3ヶ月で0.5UP）")
    rec["notes"]           = ask("その他メモ")

    return rec


# -----------------------------------------------------------------
# Excel出力
# -----------------------------------------------------------------

COLUMNS = [
    ("ID",              lambda r: r["id"]),
    ("名前",            lambda r: r["name"]),
    ("相談日",          lambda r: r["date_consulted"]),
    ("受験タイプ",      lambda r: r["exam_type"]),
    ("受験目的",        lambda r: r["exam_purpose"]),
    ("目標校・国",      lambda r: r["target_school_country"]),
    ("締め切り",        lambda r: r["deadline"]),
    # 現在スコア
    ("現スコア_総合",   lambda r: r["current_score"].get("overall", "")),
    ("現スコア_L",      lambda r: r["current_score"].get("listening", "")),
    ("現スコア_R",      lambda r: r["current_score"].get("reading", "")),
    ("現スコア_W",      lambda r: r["current_score"].get("writing", "")),
    ("現スコア_S",      lambda r: r["current_score"].get("speaking", "")),
    # 目標スコア
    ("目標スコア_総合", lambda r: r["target_score"].get("overall", "")),
    ("目標スコア_L",    lambda r: r["target_score"].get("listening", "")),
    ("目標スコア_R",    lambda r: r["target_score"].get("reading", "")),
    ("目標スコア_W",    lambda r: r["target_score"].get("writing", "")),
    ("目標スコア_S",    lambda r: r["target_score"].get("speaking", "")),
    # 受験履歴・予定
    ("前回受験日",      lambda r: r["last_exam_date"]),
    ("次回受験予定",    lambda r: r["next_exam_date"]),
    # 学習状況
    ("現在の学習状況",  lambda r: r["study_situation"]),
    ("弱点・課題",      lambda r: r["weak_points"]),
    ("不安・悩み",      lambda r: r["concerns"]),
    ("スケジュール計画",lambda r: r["schedule_plan"]),
    ("メモ",            lambda r: r["notes"]),
]


def export_excel(records: list) -> Path:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("\n  openpyxl が必要です。pip install openpyxl を実行してください。")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "個別相談一覧"

    # ヘッダースタイル
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # グループヘッダー（2行目のカラーグループ）
    group_fills = {
        "基本情報":   PatternFill("solid", fgColor="2E75B6"),
        "現在スコア": PatternFill("solid", fgColor="70AD47"),
        "目標スコア": PatternFill("solid", fgColor="ED7D31"),
        "受験日程":   PatternFill("solid", fgColor="7030A0"),
        "学習状況":   PatternFill("solid", fgColor="C00000"),
    }

    col_groups = [
        ("基本情報",   7),
        ("現在スコア", 5),
        ("目標スコア", 5),
        ("受験日程",   2),
        ("学習状況",   4),
    ]

    col = 1
    for group_name, span in col_groups:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
        cell = ws.cell(row=1, column=col, value=group_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = group_fills[group_name]
        cell.alignment = center
        cell.border = border
        col += span

    # カラムヘッダー（2行目）
    for i, (name, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=2, column=i, value=name)
        cell.font = header_font
        cell.fill = PatternFill("solid", fgColor="2F5496")
        cell.alignment = center
        cell.border = border

    # データ行
    for row_idx, rec in enumerate(records, start=3):
        for col_idx, (_, getter) in enumerate(COLUMNS, start=1):
            val = getter(rec)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border
            cell.alignment = left if col_idx > 7 else center
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="EBF3FB")

    # 列幅調整
    widths = [5, 12, 12, 12, 20, 18, 12,
              8, 6, 6, 6, 6,
              8, 6, 6, 6, 6,
              12, 12,
              30, 25, 30, 30, 25]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A3"

    wb.save(EXCEL_OUTPUT)
    return EXCEL_OUTPUT


# -----------------------------------------------------------------
# メイン
# -----------------------------------------------------------------

def main():
    records = load_consultations()

    print()
    print("=" * 46)
    print("  個別相談トラッカー")
    print("=" * 46)
    print(f"  登録済み相談者: {len(records)} 名")
    print()
    print("  1. 新規相談者を追加する")
    print("  2. Excelに出力する")
    print("  3. 一覧を表示する")
    print("  4. 戻る")
    print()

    choice = input("番号を入力してください (1〜4): ").strip()

    if choice == "1":
        rec = input_new_record(records)
        records.append(rec)
        save_consultations(records)
        print(f"\n  ID:{rec['id']} {rec['name']} を保存しました。")

        out = export_excel(records)
        if out:
            print(f"  Excelも更新しました: {out.name}")

    elif choice == "2":
        if not records:
            print("  まだ相談者データがありません。")
            return
        out = export_excel(records)
        if out:
            print(f"\n  Excelを出力しました: {out}")

    elif choice == "3":
        if not records:
            print("  まだ相談者データがありません。")
            return
        print()
        print(f"  {'ID':>3}  {'名前':<12}  {'相談日':<12}  {'現スコア':>8}  {'目標':>6}  {'次回受験':<12}  目的")
        print("  " + "-" * 75)
        for r in records:
            print(
                f"  {r['id']:>3}  {r['name']:<12}  {r['date_consulted']:<12}"
                f"  {r['current_score'].get('overall', '-'):>8}"
                f"  {r['target_score'].get('overall', '-'):>6}"
                f"  {r['next_exam_date']:<12}  {r['exam_purpose']}"
            )
        print()


if __name__ == "__main__":
    main()
