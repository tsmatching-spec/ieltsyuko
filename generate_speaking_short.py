#!/usr/bin/env python3
"""
IELTS Speaking Short Video Script Generator

Generates short video scripts (ショート動画) for IELTS Speaking practice
from three perspectives: 採点官 (Examiner), 受験生 (Candidate), コーチ (Coach)

Usage:
    python generate_speaking_short.py              # auto-select next topic
    python generate_speaking_short.py --id 3       # specify topic by ID
    python generate_speaking_short.py --list       # list all topics
    python generate_speaking_short.py --reset      # reset all topics to undone
    python generate_speaking_short.py --custom "Your own question here"
"""

import json
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

_env_file = Path(__file__).parent / ".env"
if _env_file.exists():
    for line in _env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

import anthropic

TOPICS_FILE = Path(__file__).parent / "speaking_topics.json"
SCRIPTS_DIR = Path(__file__).parent / "scripts" / "speaking_shorts"


def load_topics():
    with open(TOPICS_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_topics(data):
    with open(TOPICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def list_topics(data):
    print("\n=== IELTS Speaking Short Video Topic List ===\n")
    for t in data["topics"]:
        status = "✓" if t["done"] else "○"
        print(f"[{status}] ID {t['id']:2d} | Part {t['part']} | {t['theme']:<20} | {t['question'][:50]}...")
    print()


def select_topic(data, topic_id=None):
    topics = data["topics"]
    if topic_id:
        matches = [t for t in topics if t["id"] == topic_id]
        if not matches:
            print(f"Error: Topic ID {topic_id} not found.")
            sys.exit(1)
        return matches[0]
    for t in topics:
        if not t["done"]:
            return t
    print("All topics completed! Use --reset to start over.")
    sys.exit(0)


def build_prompt(topic):
    question = topic["question"]
    part = topic.get("part", 1)
    theme = topic.get("theme", "")
    grammar_focus = topic.get("grammar_focus", "")

    return f"""あなたはIELTSスピーキング指導のプロで、YouTubeチャンネル「IELTSゆうこ」のショート動画台本ライターです。

以下のIELTSスピーキングの質問について、3つの視点（採点官・受験生・コーチ）を使ったショート動画の台本を作成してください。

**質問情報**
- Part: {part}
- テーマ: {theme}
- 質問: {question}
- 文法フォーカス: {grammar_focus}

---

**動画コンセプト**

このショート動画は以下の流れで構成されます：
1. 採点官が「何を見ているか」をひとことで示す
2. 受験生が最初の（よくある）答えを言う
3. コーチが改善のヒントを教える
4. 受験生が中級レベルの答えを言う（ワンポイント改善版）
5. コーチがさらに上を目指すテクニックを教える
6. 受験生が上級レベルの答えを言う（文法・語彙を駆使した版）
7. 採点官がスコアとポイントをひとことでまとめる

---

**出力フォーマット（Markdown形式）**

# [動画タイトル（例：「一人で勉強する派？グループ派？ → IELTSで差がつく答え方3段階」）]

## 📋 動画概要
- **対象**: IELTS Speaking Part {part}
- **テーマ**: {theme}
- **文法フォーカス**: {grammar_focus}
- **想定時間**: 45〜60秒

---

## 🎬 台本

### オープニング（0〜5秒）
> 採点官・受験生・コーチが一言ずつ自己紹介する短いカット

**【採点官】**:（採点官として視聴者に語りかける一言。何を見ているか）

**【コーチ】**:（今日のテーマと学ぶポイントを一言で紹介）

---

### 本編（5〜50秒）

**【コーチ】**: 今日の質問はこちら！

> 💬 質問テロップ: "{question}"

**【受験生（初級）】**:（ありがちな短くシンプルな答え。文法的には正しいが内容が浅い）

**【採点官（心の声）】**:（採点官として、この答えに感じる物足りなさをひとことで）

**【コーチ】**:（改善ポイントを一言。「〇〇を使うと自然に差がつきます！」のスタイル）

> 📌 テクニックテロップ: （コーチが教えたテクニックの名前・キーワード）

**【受験生（中級）】**:（コーチのアドバイスを活かした中級レベルの答え）

**【コーチ】**:（さらに上のレベルへのヒント。文法・語彙のテクニックを具体的に）

> 📌 テクニックテロップ: （上級テクニックの名前・キーワード）

**【受験生（上級）】**:（{grammar_focus}を使った上級レベルの答え。自然で印象的）

---

### エンディング（50〜60秒）

**【採点官】**:（上級の答えへの評価。「Band 7以上ねらえます！」など具体的なスコア感）

**【コーチ】**:（視聴者へのひとこと締め。「試してみて！」「保存して練習しよう！」など）

---

## 📝 スクリプト解説

### 初級 → 中級 → 上級の変化ポイント

| レベル | 答えの例 | 使っているテクニック |
|--------|---------|---------------------|
| 初級 | （初級の答えを再掲） | シンプルな現在形のみ |
| 中級 | （中級の答えを再掲） | （使ったテクニック） |
| 上級 | （上級の答えを再掲） | {grammar_focus} |

### コーチのワンポイント解説
（中級・上級で使ったテクニックを日本語で100字以内で解説）

### 採点官が見ているポイント（Part {part}）
（Part {part}で採点官が特に重視する3つのポイントを箇条書きで）

---

## 🎨 撮影・編集メモ
- **画面構成**: 採点官（右上）・受験生（中央）・コーチ（左下）の3分割、またはスワイプ切替
- **テロップスタイル**: 質問は大きく白抜き、テクニック名はオレンジ強調
- **BGM**: 軽快でテンポよい学習系BGM
- **字幕**: 全セリフに日本語字幕（英語発話部分は英語+日本語訳）
- **エンド画面**: 「保存して練習！」CTAと関連動画サムネ
"""


def generate_script(topic):
    client = anthropic.Anthropic()

    print(f"\nGenerating speaking short script for: [{topic.get('theme', 'Custom')}] Part {topic.get('part', '-')}")
    print(f"Question: {topic['question'][:60]}...")
    print("Calling Claude API...\n")

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4096,
        messages=[
            {"role": "user", "content": build_prompt(topic)}
        ]
    )

    return message.content[0].text


def save_script(topic, script_text):
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    topic_id = topic.get("id", "custom")
    theme = topic.get("theme", "custom").replace(" ", "_").replace("/", "-")
    filename = f"{date_str}_speaking_part{topic.get('part', 1)}_id{topic_id}_{theme}.md"
    filepath = SCRIPTS_DIR / filename

    header = f"""---
generated: {datetime.now().isoformat()}
type: speaking_short
topic_id: {topic_id}
part: {topic.get('part', 1)}
theme: {topic.get('theme', 'Custom')}
grammar_focus: {topic.get('grammar_focus', '')}
question: "{topic['question']}"
---

"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(header + script_text)

    return filepath


def main():
    parser = argparse.ArgumentParser(description="IELTS Speaking Short Video Script Generator")
    parser.add_argument("--id", type=int, help="Specify topic ID")
    parser.add_argument("--list", action="store_true", help="List all topics")
    parser.add_argument("--reset", action="store_true", help="Reset all topics to undone")
    parser.add_argument("--custom", metavar="QUESTION", help="Use a custom question")
    parser.add_argument("--part", type=int, default=1, help="IELTS Speaking Part (1/2/3), used with --custom")
    parser.add_argument("--grammar", metavar="FOCUS", help="Grammar focus, used with --custom")
    args = parser.parse_args()

    if args.custom:
        topic = {
            "id": "custom",
            "part": args.part,
            "theme": "Custom",
            "question": args.custom,
            "grammar_focus": args.grammar or "Mixed Conditional / Advanced vocabulary",
            "done": False,
        }
        script = generate_script(topic)
        filepath = save_script(topic, script)
        print(f"\nScript saved to: {filepath}")
        print("Done!")
        return

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

    topic = select_topic(data, args.id)
    print(f"\n--- Selected Topic ---")
    print(f"ID     : {topic['id']}")
    print(f"Part   : {topic['part']}")
    print(f"Theme  : {topic['theme']}")
    print(f"Grammar: {topic['grammar_focus']}")
    print(f"Question: {topic['question'][:80]}...")

    script = generate_script(topic)
    filepath = save_script(topic, script)

    for t in data["topics"]:
        if t["id"] == topic["id"]:
            t["done"] = True
    save_topics(data)

    print(f"\nScript saved to: {filepath}")
    print("Done!")


if __name__ == "__main__":
    main()
