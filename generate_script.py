#!/usr/bin/env python3
"""
IELTS Writing Task 2 YouTube Script Generator

Usage:
    python generate_script.py              # auto-select next topic
    python generate_script.py --id 3       # specify topic by ID
    python generate_script.py --list       # list all topics
    python generate_script.py --reset      # reset all topics to undone
"""

import json
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

# load .env if present
_env_file = Path(__file__).parent / ".env"
if _env_file.exists():
    for line in _env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

import anthropic

TOPICS_FILE = Path(__file__).parent / "topics.json"
SCRIPTS_DIR = Path(__file__).parent / "scripts"
RESEARCH_DIR = Path(__file__).parent / "research"


def load_topics():
    with open(TOPICS_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_topics(data):
    with open(TOPICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def list_topics(data):
    print("\n=== IELTS Writing Task 2 Topic List ===\n")
    for t in data["topics"]:
        status = "✓" if t["done"] else "○"
        print(f"[{status}] ID {t['id']:2d} | {t['theme']:<25} | {t['essay_type']}")
    print()


def select_topic(data, topic_id=None):
    topics = data["topics"]
    if topic_id:
        matches = [t for t in topics if t["id"] == topic_id]
        if not matches:
            print(f"Error: Topic ID {topic_id} not found.")
            sys.exit(1)
        return matches[0]
    # auto-select: first undone topic
    for t in topics:
        if not t["done"]:
            return t
    print("All topics have been completed! Use --reset to start over.")
    sys.exit(0)


def load_latest_research() -> str | None:
    """Return the text of the most recent SEO research report, if any."""
    if not RESEARCH_DIR.exists():
        return None
    reports = sorted(RESEARCH_DIR.glob("*_seo_report.md"), reverse=True)
    if not reports:
        return None
    return reports[0].read_text(encoding="utf-8")


def build_prompt(topic, research_text: str | None = None):
    seo_context = ""
    if research_text:
        # extract the recommendations section to keep the prompt concise
        marker = "## 4. 勝てる動画タイトル・フォーマット"
        if marker in research_text:
            seo_context = f"""
## SEOリサーチ情報（最新レポートより）
以下のリサーチ結果を参考に、タイトル・説明文・タグ・サムネイルをSEO最適化してください。

{research_text[research_text.index(marker):]}

---
"""
    return f"""あなたはIELTS指導のプロで、YouTubeチャンネル「IELTSゆうこ」の台本ライターです。
以下のIELTS Writing Task 2のトピックについて、中級者（スコア5.5〜6.5を目指す学習者）向けの
YouTubeスクリプトを日本語で作成してください。
{seo_context}
**動画の条件**
- 長さ: 15〜20分（読み上げ速度で約3,000〜4,000文字のナレーション）
- 対象: 中級者（Band 5.5〜6.5目標）
- 言語: 日本語（英語例文は英語のまま、その後日本語で解説）
- スタイル: 親しみやすく、わかりやすい解説。テンポよく、視聴者を飽きさせない

**トピック情報**
- テーマ: {topic['theme']}
- エッセイタイプ: {topic['essay_type']}
- 問題文: {topic['question']}

---

**出力フォーマット（Markdown形式）**

# [動画タイトル（SEOを意識した日本語タイトル、60文字以内）]

## SEO情報
- **YouTube説明文**: （200文字程度。検索キーワードを自然に含める）
- **タグ**: （カンマ区切りで15個）
- **サムネイル案**: （視覚的に目を引くサムネイルのコンセプトを1〜2文で説明）

---

## 動画構成（タイムライン）
| 時間 | セクション | 内容 |
|------|-----------|------|
（各セクションの目安時間と内容を記載）

---

## スクリプト本文

### 【オープニング】（約1分）
（挨拶、自己紹介、今日のテーマ紹介。視聴者の興味を引くフック）

### 【問題文の確認と読み解き】（約2分）
（問題文を一緒に読み、何を書くべきかを明確にする）

### 【エッセイ構成の解説】（約2分）
（{topic['essay_type']}のエッセイ構成テンプレートを説明）

### 【Introduction の書き方】（約3分）
（パラフレーズのテクニック、Thesis statementの書き方を具体例付きで解説）
- サンプル Introduction（英語）:
  （サンプルを記載）
- 解説:

### 【Body Paragraph の書き方】（約5分）
（各ボディパラグラフの構成、トピックセンテンス、支持文、例文の書き方）
- Body 1 サンプル（英語）:
  （サンプルを記載）
- 解説:
- Body 2 サンプル（英語）:
  （サンプルを記載）
- 解説:

### 【Conclusion の書き方】（約2分）
（まとめ方、意見の再提示）
- サンプル Conclusion（英語）:
  （サンプルを記載）
- 解説:

### 【スコアアップの必須語彙・表現】（約3分）
（このテーマ・エッセイタイプで使える高スコア語彙・フレーズを5〜8個、使用例付きで紹介）

| 表現 | 意味 | 使用例 |
|------|------|--------|
（表を埋める）

### 【よくある間違いと改善ポイント】（約2分）
（中級者がやりがちなミスを3点挙げ、改善例を示す）

### 【エンディング】（約1分）
（まとめ、次回予告、チャンネル登録・高評価のお願い）

---

## フルサンプルエッセイ（Band 6.5レベル）

（上記問題文に対するフルエッセイ。約250〜280語）

---

## 撮影・編集メモ
- BGM:
- テロップ挿入箇所:
- 画面切り替えポイント:
"""


def generate_script(topic, research_text: str | None = None):
    client = anthropic.Anthropic()

    print(f"\nGenerating script for: [{topic['theme']}] {topic['essay_type']}")
    print(f"Topic ID: {topic['id']}")
    if research_text:
        print("SEO research context: loaded")
    print("Calling Claude API...\n")

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8096,
        messages=[
            {"role": "user", "content": build_prompt(topic, research_text)}
        ]
    )

    return message.content[0].text


def save_script(topic, script_text):
    SCRIPTS_DIR.mkdir(exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{date_str}_id{topic['id']:02d}_{topic['theme'].replace(' ', '_').replace('&', 'and')}.md"
    filepath = SCRIPTS_DIR / filename

    header = f"""---
generated: {datetime.now().isoformat()}
topic_id: {topic['id']}
theme: {topic['theme']}
essay_type: {topic['essay_type']}
question: "{topic['question']}"
---

"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(header + script_text)

    return filepath


def main():
    parser = argparse.ArgumentParser(description="IELTS Writing Task 2 YouTube Script Generator")
    parser.add_argument("--id", type=int, help="Specify topic ID")
    parser.add_argument("--list", action="store_true", help="List all topics")
    parser.add_argument("--reset", action="store_true", help="Reset all topics to undone")
    parser.add_argument("--research", metavar="FILE",
                        help="SEO research report to use (default: latest in research/)")
    parser.add_argument("--no-research", action="store_true",
                        help="Ignore SEO research even if a report exists")
    args = parser.parse_args()

    data = load_topics()

    if args.list:
        list_topics(data)
        return

    if args.reset:
        for t in data["topics"]:
            t["done"] = False
        save_topics(data)
        print("All topics have been reset.")
        return

    # load SEO research context
    research_text = None
    if not args.no_research:
        if args.research:
            research_path = Path(args.research)
            if research_path.exists():
                research_text = research_path.read_text(encoding="utf-8")
                print(f"SEO research: {research_path}")
            else:
                print(f"Warning: research file not found: {args.research}")
        else:
            research_text = load_latest_research()
            if research_text:
                reports = sorted(RESEARCH_DIR.glob("*_seo_report.md"), reverse=True)
                print(f"SEO research: {reports[0]} (latest)")

    topic = select_topic(data, args.id)
    print(f"\n--- Selected Topic ---")
    print(f"ID     : {topic['id']}")
    print(f"Theme  : {topic['theme']}")
    print(f"Type   : {topic['essay_type']}")
    print(f"Question: {topic['question'][:80]}...")

    script = generate_script(topic, research_text)
    filepath = save_script(topic, script)

    # mark as done
    for t in data["topics"]:
        if t["id"] == topic["id"]:
            t["done"] = True
    save_topics(data)

    print(f"\nScript saved to: {filepath}")
    print("Done!")


if __name__ == "__main__":
    main()
