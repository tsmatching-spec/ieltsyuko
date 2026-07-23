#!/usr/bin/env python3
"""ちひろさんの添削フィードバック（気をつけるポイント）を管理する共通モジュール"""

from datetime import datetime
from pathlib import Path

FEEDBACK_FILE = Path(__file__).parent / "mentor_feedback.md"
CHECKLIST_HEADER = "## 気をつけるポイント"


def load_feedback_points() -> str | None:
    """蓄積された「気をつけるポイント」チェックリストを返す（未登録なら None）。"""
    if not FEEDBACK_FILE.exists():
        return None
    text = FEEDBACK_FILE.read_text(encoding="utf-8")
    if CHECKLIST_HEADER not in text:
        return None
    checklist = text.split(CHECKLIST_HEADER, 1)[1].strip()
    return checklist or None


def add_feedback_point(point: str, today: str | None = None) -> None:
    """ちひろさんの添削から得た「気をつけるポイント」を1行追加する。"""
    today = today or datetime.now().strftime("%Y-%m-%d")
    line = f"- [{today}] {point}"

    if not FEEDBACK_FILE.exists():
        FEEDBACK_FILE.write_text(
            "# メンターちひろさんフィードバック — 台本作成で気をつけるポイント\n\n"
            f"{CHECKLIST_HEADER}\n{line}\n",
            encoding="utf-8",
        )
        return

    text = FEEDBACK_FILE.read_text(encoding="utf-8")
    if CHECKLIST_HEADER in text:
        text = text.rstrip("\n") + f"\n{line}\n"
    else:
        text = text.rstrip("\n") + f"\n\n{CHECKLIST_HEADER}\n{line}\n"
    FEEDBACK_FILE.write_text(text, encoding="utf-8")
