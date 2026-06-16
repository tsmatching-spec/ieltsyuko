#!/usr/bin/env python3
"""
IELTS Speaking フレーズ言い換えショート動画台本ジェネレーター

構成:
  ① NG例を見せる
  ② 問題提起（なぜNGか一言）
  ③ 代替表現リスト（3〜6個）
  ④ 使い方ヒント（一言締め）

Usage:
    python generate_phrase_short.py --ng "That's an interesting question"
    python generate_phrase_short.py  # インタラクティブモード
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

SCRIPTS_DIR = Path(__file__).parent / "scripts" / "phrase_shorts"


def build_prompt(ng_phrase: str, context: str = "") -> str:
    context_section = f"\n追加コンテキスト: {context}" if context else ""
    return f"""あなたはIELTSスピーキング指導のプロで、YouTubeチャンネル「IELTSゆうこ」のショート動画台本ライターです。

以下のNG表現について、ショート動画台本を作成してください。{context_section}

**NG表現**: "{ng_phrase}"

---

**動画の構成（この順番で、シンプルに）**:
① NG例を視聴者に見せる（このフレーズを使っている様子）
② 問題提起（なぜNGなのか、一言で）
③ 代替表現リスト（3〜6個、場面別に）
④ 使い方ヒント（一言締め）

---

**出力フォーマット（Markdown形式）**

# [タイトル例: 「{ng_phrase}」はもう卒業！代わりに使える〇〇フレーズ]

## 📋 動画概要
- **対象**: IELTS Speaking 全パート
- **NG表現**: {ng_phrase}
- **NG理由**: （採点官視点で、なぜこれを多用するとスコアが下がるか一言）
- **想定時間**: 30〜45秒

---

## 🎬 台本

### ① オープニング／NG例（0〜8秒）

> 🎙️ 画面テロップ（大きく）: "{ng_phrase}"

**ナレーション（ゆうこ）**: （このフレーズが試験でどう使われるか、共感を引く一言。「これ言いがちじゃないですか？」スタイル）

---

### ② 問題提起（8〜15秒）

**ナレーション**: （なぜNGか。採点官の視点から一言。「実は〇〇なんです」スタイル）

> 📌 テロップ: （NG理由のキーワード、例「採点官にはこう聞こえてます」）

---

### ③ 代替表現リスト（15〜35秒）

**ナレーション**: 代わりにこれを使って！

> ✅ 代替表現テロップ（一つずつ表示）:
>
> ① [表現1]（使う場面・ニュアンスの補足）
> ② [表現2]（使う場面・ニュアンスの補足）
> ③ [表現3]（使う場面・ニュアンスの補足）
> ④ [表現4]（使う場面・ニュアンスの補足）
> ⑤ [表現5]（使う場面・ニュアンスの補足）
> ⑥ [表現6]（あれば）

---

### ④ 締め／使い方ヒント（35〜45秒）

**ナレーション**: （使い分けのコツを一言。「場面に合わせて選んで、繰り返さないのがポイント！」スタイル）

> 💾 エンドテロップ: 「保存して練習しよう！」

---

## 📝 表現解説

### 代替表現まとめ

| 表現 | 使うシーン | ニュアンス |
|------|----------|----------|
（各表現を表にまとめる）

### なぜ "{ng_phrase}" を避けるべきか
（採点官の視点から、100字以内で解説）

---

## 🎨 撮影・編集メモ
- **NG表現**: 赤テキスト or バツ印で強調
- **代替表現**: 緑テキスト or チェックマークで一つずつ表示
- **BGM**: テンポよい学習系BGM（代替表現が出るたびにポップな効果音）
- **字幕**: 全ナレーションに日本語字幕
- **エンド画面**: 「保存して練習！」CTA
"""


def generate_script(ng_phrase: str, context: str = "") -> str:
    client = anthropic.Anthropic()
    print(f"\nNG表現「{ng_phrase}」の台本を生成中...")

    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=3000,
        messages=[
            {"role": "user", "content": build_prompt(ng_phrase, context)}
        ]
    )
    return message.content[0].text


def save_script(script_text: str, ng_phrase: str) -> Path:
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = ng_phrase[:30].replace(" ", "_").replace("'", "").replace(".", "").replace("?", "")
    filename = f"{date_str}_phrase_{safe_name}.md"
    filepath = SCRIPTS_DIR / filename

    header = f"""---
generated: {datetime.now().isoformat()}
type: phrase_short
ng_phrase: "{ng_phrase}"
---

"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(header + script_text)
    return filepath


def main():
    parser = argparse.ArgumentParser(
        description="IELTS Speaking フレーズ言い換えショート動画台本生成"
    )
    parser.add_argument("--ng", metavar="PHRASE", help="NG表現（例: 'That's an interesting question'）")
    parser.add_argument("--context", metavar="TEXT", default="", help="追加コンテキスト（任意）")
    args = parser.parse_args()

    if args.ng:
        ng_phrase = args.ng
        context = args.context
    else:
        print()
        print("【フレーズ言い換えショート台本生成】")
        print()
        ng_phrase = input("NG表現を入力してください: ").strip().strip('"')
        if not ng_phrase:
            print("キャンセルしました。")
            sys.exit(0)
        context = input("追加コンテキスト（なければEnter）: ").strip()

    script_text = generate_script(ng_phrase, context)
    filepath = save_script(script_text, ng_phrase)

    print(f"\n台本を保存しました: {filepath.relative_to(Path(__file__).parent)}")
    print("Done!")


if __name__ == "__main__":
    main()
