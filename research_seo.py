#!/usr/bin/env python3
"""
IELTS Writing Task 2 YouTube SEO Research Tool

YouTubeでIELTSライティング動画を伸ばすための企画リサーチを行い、
高需要・低競合のトピックをランキング形式でレポート出力します。

Usage:
    python research_seo.py                  # フルリサーチ実行
    python research_seo.py --add-topics     # リサーチ結果をtopics.jsonに追加
    python research_seo.py --quick          # クイックリサーチ（主要キーワードのみ）
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
RESEARCH_DIR = Path(__file__).parent / "research"

# -----------------------------------------------------------------
# Web search queries to gather SEO signals
# -----------------------------------------------------------------

SEARCH_QUERIES = [
    "IELTS ライティング Task 2 YouTube 人気",
    "IELTS writing task 2 tips Japanese YouTube",
    "IELTSライティング 勉強法 youtube 2024",
    "IELTS band 7 writing task 2 youtube",
    "IELTSライティング 独学 おすすめ",
]

COMPETITOR_QUERIES = [
    "IELTSライティング解説 youtube チャンネル",
    "IELTS writing japanese youtube channel",
    "IELTS 日本語 youtube 人気動画",
]


# -----------------------------------------------------------------
# Prompts
# -----------------------------------------------------------------

def build_search_synthesis_prompt(search_results: list[dict]) -> str:
    results_text = "\n\n".join(
        f"[検索クエリ: {r['query']}]\n{r['result']}"
        for r in search_results
    )
    return f"""あなたはYouTube SEOとIELTS教育コンテンツの専門家です。
以下のウェブ検索結果をもとに、「IELTSゆうこ」チャンネル（IELTS Writing Task 2特化、中級者向け日本語チャンネル）の
YouTube SEOリサーチレポートを作成してください。

## 検索結果データ
{results_text}

---

以下の構成でMarkdown形式のレポートを作成してください：

# IELTS Writing Task 2 YouTube SEO リサーチレポート
**作成日**: {datetime.now().strftime('%Y年%m月%d日')}

---

## 1. 市場概況
（IELTSライティング関連のYouTubeコンテンツ市場の現状。競合チャンネルの傾向、視聴者ニーズのまとめ）

---

## 2. 高需要キーワード分析

### 2-1. 検索ボリューム推定トップキーワード
| キーワード | 推定ニーズ | 競合レベル | 優先度 |
|-----------|-----------|-----------|-------|
（10〜15個のキーワードを記載。優先度は ◎/○/△ で評価）

### 2-2. 狙い目の長尾キーワード（ニッチ・低競合）
（競合が少なく、検索意図が明確なキーワードを5〜8個リストアップ）

---

## 3. 競合チャンネル分析

### 主要競合チャンネル
（判明した競合チャンネルの特徴・コンテンツ傾向・強みを整理）

### コンテンツギャップ（競合が扱っていない／不足しているテーマ）
（「IELTSゆうこ」が差別化できる空白領域を箇条書きで）

---

## 4. 勝てる動画タイトル・フォーマット

### 高クリック率が期待できるタイトルパターン
（YouTubeで日本語IELTS視聴者にウケやすいタイトル構成を5パターン例示）

例:
1. `【IELTS】○○で絶対NG！〜〜の正しい書き方【Task 2】`
2. ...

### サムネイル戦略
（クリックを取りやすいサムネイルのデザイン・コピーの方向性）

---

## 5. 推奨コンテンツカレンダー（次の6本）

以下の形式で、優先度順に6本の動画企画を提案してください。
各企画はSEOリサーチに基づき、検索需要・競合状況・視聴者ニーズを考慮すること。

### 動画企画 #1（最優先）
- **テーマ**:
- **エッセイタイプ**: （Opinion / Discussion+Opinion / Problem-Solution / Cause-Effect）
- **想定問題文**: （実際のIELTS Task 2形式で記載）
- **ターゲットキーワード**: （メイン1個 + サブ2〜3個）
- **推奨タイトル**: （60文字以内、SEO最適化済み）
- **企画理由**: （なぜ今この動画が伸びるか、2〜3文で）
- **差別化ポイント**: （競合と何が違うか）

### 動画企画 #2
（同形式で）

### 動画企画 #3
（同形式で）

### 動画企画 #4
（同形式で）

### 動画企画 #5
（同形式で）

### 動画企画 #6
（同形式で）

---

## 6. チャンネル成長戦略メモ
（短期・中期で取り組むべきSEO・コンテンツ戦略を箇条書きで。投稿頻度、プレイリスト構成、概要欄の書き方など実践的なヒントを含む）
"""


def build_topics_extraction_prompt(report_text: str, existing_ids: list[int]) -> str:
    next_id = max(existing_ids) + 1 if existing_ids else 1
    return f"""以下はIELTS YouTube SEOリサーチレポートです。
レポート内の「推奨コンテンツカレンダー」セクションから、動画企画データをJSON配列として抽出してください。

## レポート
{report_text}

---

以下のJSON形式で出力してください（コードブロックなし、JSONのみ）：

[
  {{
    "id": {next_id},
    "theme": "テーマ名（英語）",
    "question": "想定問題文（英語）",
    "essay_type": "Opinion|Discussion + Opinion|Problem-Solution|Cause-Effect",
    "seo_keywords": ["キーワード1", "キーワード2", "キーワード3"],
    "recommended_title": "推奨YouTube動画タイトル（日本語）",
    "priority": 1,
    "source": "seo_research",
    "difficulty": "intermediate",
    "done": false
  }},
  ...
]

ID は {next_id} から連番で振ること。企画数分（最大6件）のオブジェクトを返すこと。
"""


# -----------------------------------------------------------------
# Core functions
# -----------------------------------------------------------------

def run_web_searches(client: anthropic.Anthropic, queries: list[str]) -> list[dict]:
    """Run web searches using Claude's web_search tool."""
    results = []
    for query in queries:
        print(f"  Searching: {query}")
        try:
            response = client.messages.create(
                model="claude-opus-4-5",
                max_tokens=1024,
                tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 1}],
                messages=[{
                    "role": "user",
                    "content": f"次のクエリで検索し、見つかった情報を日本語で要約してください: {query}"
                }]
            )
            # extract text from response
            text_parts = [
                block.text for block in response.content
                if hasattr(block, "text")
            ]
            results.append({"query": query, "result": "\n".join(text_parts) or "(結果なし)"})
        except Exception as e:
            print(f"    Warning: search failed for '{query}': {e}")
            results.append({"query": query, "result": f"(検索エラー: {e})"})
    return results


def synthesize_report(client: anthropic.Anthropic, search_results: list[dict]) -> str:
    print("Synthesizing SEO report with Claude...")
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8096,
        messages=[{
            "role": "user",
            "content": build_search_synthesis_prompt(search_results)
        }]
    )
    return response.content[0].text


def extract_topics_from_report(client: anthropic.Anthropic, report_text: str, existing_ids: list[int]) -> list[dict]:
    print("Extracting topic recommendations from report...")
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": build_topics_extraction_prompt(report_text, existing_ids)
        }]
    )
    raw = response.content[0].text.strip()
    # strip possible markdown fences
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw)


def save_report(report_text: str) -> Path:
    RESEARCH_DIR.mkdir(exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    filepath = RESEARCH_DIR / f"{date_str}_seo_report.md"

    header = f"""---
generated: {datetime.now().isoformat()}
type: seo_research
channel: IELTSゆうこ
target: IELTS Writing Task 2（中級者向け）
---

"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(header + report_text)
    return filepath


def add_topics_to_json(new_topics: list[dict]) -> int:
    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)

    existing_ids = {t["id"] for t in data["topics"]}
    added = 0
    for t in new_topics:
        if t["id"] not in existing_ids:
            data["topics"].append(t)
            added += 1

    with open(TOPICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return added


def get_existing_ids() -> list[int]:
    if not TOPICS_FILE.exists():
        return []
    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)
    return [t["id"] for t in data["topics"]]


# -----------------------------------------------------------------
# Main
# -----------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="IELTS Writing YouTube SEO Research Tool")
    parser.add_argument("--add-topics", action="store_true",
                        help="Add recommended topics from research to topics.json")
    parser.add_argument("--quick", action="store_true",
                        help="Quick mode: skip competitor queries, use fewer searches")
    args = parser.parse_args()

    client = anthropic.Anthropic()

    queries = SEARCH_QUERIES if args.quick else SEARCH_QUERIES + COMPETITOR_QUERIES
    print(f"\n=== IELTS Writing YouTube SEO Research ===")
    print(f"Mode: {'Quick' if args.quick else 'Full'} ({len(queries)} searches)\n")

    print("Step 1: Web search...")
    search_results = run_web_searches(client, queries)

    print("\nStep 2: Generating SEO report...")
    report_text = synthesize_report(client, search_results)

    report_path = save_report(report_text)
    print(f"\nReport saved to: {report_path}")

    if args.add_topics:
        print("\nStep 3: Extracting topic recommendations...")
        existing_ids = get_existing_ids()
        try:
            new_topics = extract_topics_from_report(client, report_text, existing_ids)
            added = add_topics_to_json(new_topics)
            print(f"Added {added} new topics to topics.json")
        except Exception as e:
            print(f"Warning: Could not extract topics automatically: {e}")
            print("Please check the report and add topics manually.")

    print("\nDone! Next step:")
    print(f"  python generate_script.py --research {report_path}")


if __name__ == "__main__":
    main()
