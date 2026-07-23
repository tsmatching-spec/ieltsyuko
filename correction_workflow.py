#!/usr/bin/env python3
"""
添削ワークフロー（Claude Code チャット連携版）

流れ:
  1. .docx ファイルを2つ指定する
  2. テキストを抽出し、チャットに貼るプロンプトを出力する
  3. Claude Code チャットで台本を生成する
  4. 生成された台本をターミナルに貼り付けて保存する

起動: python correction_workflow.py
"""

import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from mentor_feedback import load_feedback_points

TOPICS_FILE = Path(__file__).parent / "topics.json"
SCRIPTS_DIR = Path(__file__).parent / "scripts"


# -----------------------------------------------------------------
# .docx テキスト抽出
# -----------------------------------------------------------------

def extract_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    text = re.sub(r"<w:br[^/]*/?>", "\n", xml)
    text = re.sub(r"</w:p>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text


# -----------------------------------------------------------------
# プロンプト生成
# -----------------------------------------------------------------

def build_chat_prompt(essay_text: str, feedback_text: str) -> str:
    mentor_points = load_feedback_points()
    mentor_context = ""
    if mentor_points:
        mentor_context = f"""
【メンターちひろさんからのフィードバック（気をつけるポイント）】
過去の添削で指摘された以下のポイントを守って台本を作成してください。

{mentor_points}
"""
    return f"""以下の添削済みエッセイをもとに、パターン①（スコアフック型）の動画台本を日本語で作成してください。

【エッセイファイル（学習者の提出物）】
{essay_text}

【フィードバックファイル（添削済み）】
{feedback_text}
{mentor_context}
台本の構成（パターン①スコアフック型）:
- 0:00  スコア発表フック（想定スコアを冒頭10秒で発表し「でも〇か所直した」と続ける）
- 0:45  問題文確認・採点基準（TR/CC/LR/GR）の紹介と各スコア
- 2:00  Before エッセイを読む（「どこが問題か探しながら見てください」）
- 4:30  修正箇所を採点基準ごとに解説（❌→✅形式、具体的な理由を説明）
- 10:30 After エッセイを見せる（Before との左右比較）
- 12:30 学習者の質問に回答（エッセイに質問が含まれている場合）
- 14:00 まとめ・CTA（今日の3ポイント＋「コメントにエッセイを貼って」）

出力に含めるもの:
- タイトル（スコアの数字を入れた60文字以内の日本語タイトル）
- SEO情報（説明文200文字・タグ15個・サムネイル案）
- タイムライン表
- スクリプト本文（各セクションの話し言葉）
- 撮影・編集メモ（テロップ挿入箇所・画面切り替えポイント）"""


# -----------------------------------------------------------------
# 台本の保存
# -----------------------------------------------------------------

def get_next_topic_id() -> int:
    if not TOPICS_FILE.exists():
        return 1
    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)
    return max((t["id"] for t in data["topics"]), default=0) + 1


def save_script(script_text: str, topic_id: int, theme: str) -> Path:
    SCRIPTS_DIR.mkdir(exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    safe_theme = re.sub(r"[^\w\-]", "_", theme)
    filename = f"{date_str}_id{topic_id:02d}_{safe_theme}.md"
    filepath = SCRIPTS_DIR / filename
    header = f"""---
generated: {datetime.now().isoformat()}
topic_id: {topic_id}
theme: {theme}
structure_pattern: "①スコアフック型"
---

"""
    filepath.write_text(header + script_text, encoding="utf-8")
    return filepath


def add_to_topics(topic_id: int, theme: str, question: str) -> None:
    if not TOPICS_FILE.exists():
        return
    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)
    # 既に存在するIDはスキップ
    if any(t["id"] == topic_id for t in data["topics"]):
        return
    data["topics"].append({
        "id": topic_id,
        "theme": theme,
        "question": question,
        "essay_type": "Discussion + Opinion",
        "seo_keywords": ["IELTSライティング添削", "IELTS essay correction"],
        "difficulty": "intermediate",
        "done": True,
        "source": "correction_video",
    })
    with open(TOPICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# -----------------------------------------------------------------
# 台本テキストの入力（複数行ペースト対応）
# -----------------------------------------------------------------

def paste_multiline(prompt_msg: str) -> str:
    print(prompt_msg)
    print("  （入力が終わったら、空行のあとに --- とだけ入力してEnterを押してください）")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "---":
            break
        lines.append(line)
    return "\n".join(lines).strip()


# -----------------------------------------------------------------
# メインフロー
# -----------------------------------------------------------------

def main() -> None:
    print()
    print("=" * 50)
    print("  添削台本ワークフロー（Claude Code チャット連携）")
    print("=" * 50)
    print()

    # ── STEP 1: ファイルパスを入力 ──────────────────────────
    print("【STEP 1】添削ファイルを指定する\n")
    print("  ヒント: エクスプローラーでファイルをShift+右クリック→「パスのコピー」")
    print()

    while True:
        essay_path = Path(input("  エッセイファイル (.docx) のパス: ").strip().strip('"'))
        if essay_path.exists():
            break
        print(f"  ファイルが見つかりません: {essay_path}")

    while True:
        feedback_path = Path(input("  フィードバックファイル (.docx) のパス: ").strip().strip('"'))
        if feedback_path.exists():
            break
        print(f"  ファイルが見つかりません: {feedback_path}")

    print("\n  ファイルを読み込み中...")
    try:
        essay_text = extract_docx(essay_path)
        feedback_text = extract_docx(feedback_path)
    except Exception as e:
        print(f"\nエラー: ファイルの読み込みに失敗しました。\n{e}")
        sys.exit(1)

    print("  読み込み完了！")

    # ── STEP 2: プロンプトを出力 ────────────────────────────
    print()
    print("=" * 50)
    print("【STEP 2】以下を Claude Code チャットにコピー＆ペーストして送信する")
    print("=" * 50)
    print()
    prompt = build_chat_prompt(essay_text, feedback_text)
    print(prompt)
    print()
    print("=" * 50)
    input("  チャットで台本が生成されたら、Enterキーを押して次に進む...")

    # ── STEP 3: 生成された台本を貼り付けて保存 ────────────
    print()
    print("【STEP 3】生成された台本をここに貼り付ける\n")
    script_text = paste_multiline("")

    if not script_text:
        print("台本が入力されませんでした。終了します。")
        sys.exit(0)

    # ── STEP 4: テーマ名を入力して保存 ─────────────────────
    print()
    print("【STEP 4】保存する\n")
    theme = input("  テーマ名（例: TV_Children_Correction）: ").strip()
    if not theme:
        theme = "Correction_Video"

    topic_id = get_next_topic_id()
    filepath = save_script(script_text, topic_id, theme)

    # topics.json に問題文を追加（任意）
    # フィードバックファイルの先頭行を問題文として使う
    question_line = next(
        (line for line in feedback_text.splitlines() if len(line) > 40 and line[0].isupper()),
        "（問題文未取得）"
    )
    add_to_topics(topic_id, theme, question_line)

    rel_path = filepath.relative_to(Path(__file__).parent)
    print()
    print(f"  台本を保存しました: {rel_path}")
    print(f"  topics.json に ID {topic_id} として追加しました。")
    print()
    print("完了！メモ帳やVSCodeでファイルを開いて確認してください。")
    print()


if __name__ == "__main__":
    main()
