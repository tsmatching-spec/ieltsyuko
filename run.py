#!/usr/bin/env python3
"""
IELTSゆうこ 動画制作ツール
起動: python run.py
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR    = Path(__file__).parent
ENV_FILE    = BASE_DIR / ".env"
TOPICS_FILE = BASE_DIR / "topics.json"
SPEAKING_TOPICS_FILE = BASE_DIR / "speaking_topics.json"
RESEARCH_DIR = BASE_DIR / "research"
SCRIPTS_DIR  = BASE_DIR / "scripts"


# -----------------------------------------------------------------
# .env ローダー
# -----------------------------------------------------------------

def load_env_file() -> None:
    if not ENV_FILE.exists():
        return
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


# -----------------------------------------------------------------
# APIキー確認（必要な機能だけで呼ぶ）
# -----------------------------------------------------------------

def ensure_api_key() -> bool:
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key and not key.startswith("sk-ant-xxx"):
        return True

    print()
    print("  この機能には Anthropic APIキーが必要です。")
    print("  取得: https://console.anthropic.com/ → API Keys")
    print()
    key = input("  APIキーを貼り付けてください（スキップは Enter）: ").strip()
    if not key:
        return False

    ENV_FILE.write_text(f"ANTHROPIC_API_KEY={key}\n", encoding="utf-8")
    os.environ["ANTHROPIC_API_KEY"] = key
    print("  .env に保存しました。\n")
    return True


# -----------------------------------------------------------------
# メニュー
# -----------------------------------------------------------------

def show_main_menu() -> str:
    print()
    print("=" * 46)
    print("  IELTSゆうこ 動画制作ツール")
    print("=" * 46)
    print()
    print("  ── Claude Code チャット連携（追加料金なし）──")
    print("  1. 添削台本を作る")
    print("  2. SEOリサーチ結果を保存する")
    print()
    print("  ── スタンドアロン実行（APIキー必要）──────")
    print("  3. SEOリサーチを自動実行する")
    print("  4. 通常台本を自動生成する")
    print("  5. スピーキングショート動画台本を生成する")
    print()
    print("  6. トピック一覧を見る")
    print("  7. 終了")
    print()
    while True:
        choice = input("番号を入力してください (1〜7): ").strip()
        if choice in ("1", "2", "3", "4", "5", "6", "7"):
            return choice
        print("  1〜7 の番号を入力してください")


# -----------------------------------------------------------------
# メニュー1: 添削台本を作る（チャット連携）
# -----------------------------------------------------------------

def run_correction_workflow() -> None:
    import correction_workflow
    correction_workflow.main()


# -----------------------------------------------------------------
# メニュー2: SEOリサーチ結果を保存する（チャット連携）
# -----------------------------------------------------------------

def save_seo_research() -> None:
    print()
    print("【SEOリサーチ結果を保存】")
    print()
    print("Claude Code チャットで「SEOリサーチして」を実行した後に使います。")
    print()

    # レポートテキストを貼り付けてもらう
    print("チャットで生成されたレポートのテキストを貼り付けてください。")
    print("（終わったら空行のあとに --- とだけ入力してEnterを押す）")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "---":
            break
        lines.append(line)
    report_text = "\n".join(lines).strip()

    if not report_text:
        print("テキストが入力されませんでした。")
        return

    # research/ に保存
    RESEARCH_DIR.mkdir(exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = RESEARCH_DIR / f"seo_report_{date_str}.md"
    report_path.write_text(report_text, encoding="utf-8")
    print(f"\nレポートを保存しました: research/seo_report_{date_str}.md")

    # topics.json に企画を追加するか確認
    add = input("\nチャットで提案された企画を topics.json に追加しますか？ (y/n): ").strip().lower()
    if add != "y":
        print("完了しました。")
        return

    print()
    print("追加したいトピックを入力してください（複数入力可）。")
    print("入力が終わったら空のテーマ名でEnterを押す。")
    print()

    if not TOPICS_FILE.exists():
        print("エラー: topics.json が見つかりません。")
        return

    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)

    next_id = max((t["id"] for t in data["topics"]), default=0) + 1
    added = 0

    while True:
        theme = input(f"  テーマ名 (ID {next_id}、空でEnterなら終了): ").strip()
        if not theme:
            break
        question = input(f"  問題文 (英語): ").strip()
        essay_type = input(f"  タイプ (Discussion+Opinion / Opinion / Problem-Solution / Cause-Effect): ").strip()
        if not essay_type:
            essay_type = "Discussion + Opinion"

        data["topics"].append({
            "id": next_id,
            "theme": theme,
            "question": question,
            "essay_type": essay_type,
            "seo_keywords": ["IELTSライティング", "IELTS Writing Task 2"],
            "difficulty": "intermediate",
            "done": False,
            "source": "seo_research",
        })
        next_id += 1
        added += 1
        print(f"  → 追加しました。\n")

    with open(TOPICS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"{added} 件の企画を topics.json に追加しました。")


# -----------------------------------------------------------------
# メニュー3: SEOリサーチ自動実行（APIキー必要）
# -----------------------------------------------------------------

def run_seo_research_auto() -> None:
    if not ensure_api_key():
        print("APIキーが設定されていないため、この機能は使えません。")
        print("メニュー「2」を使ってチャットでリサーチを実行してください。")
        return

    try:
        import research_seo
        import anthropic
    except ImportError as e:
        print(f"エラー: {e}")
        print("pip install anthropic を実行してください。")
        return

    add = input("完了後、企画を topics.json に自動追加しますか？ (y/n): ").strip().lower()
    add_topics = add == "y"

    client = anthropic.Anthropic()
    queries = research_seo.SEARCH_QUERIES + research_seo.COMPETITOR_QUERIES
    print(f"\n検索を開始します（{len(queries)} 件）...\n")

    try:
        search_results = research_seo.run_web_searches(client, queries)
        report_text = research_seo.synthesize_report(client, search_results)
        report_path = research_seo.save_report(report_text)
        print(f"レポートを保存しました: {report_path.relative_to(BASE_DIR)}")
    except Exception as e:
        print(f"\nエラー: {e}")
        return

    if add_topics:
        try:
            existing_ids = research_seo.get_existing_ids()
            new_topics = research_seo.extract_topics_from_report(client, report_text, existing_ids)
            added = research_seo.add_topics_to_json(new_topics)
            print(f"{added} 件の企画を topics.json に追加しました。")
        except Exception as e:
            print(f"自動追加に失敗しました: {e}")


# -----------------------------------------------------------------
# メニュー4: 通常台本を自動生成（APIキー必要）
# -----------------------------------------------------------------

def run_generate_script_auto() -> None:
    if not ensure_api_key():
        print("APIキーが設定されていないため、この機能は使えません。")
        return

    try:
        import generate_script
    except ImportError as e:
        print(f"エラー: {e}")
        return

    data = generate_script.load_topics()
    pending = [t for t in data["topics"] if not t["done"]]

    if not pending:
        print("すべてのトピックが完了しています。")
        return

    print()
    print("未完了のトピック:")
    for t in pending[:6]:
        print(f"  [{t['id']:2d}] {t['theme']:<30} {t['essay_type']}")
    if len(pending) > 6:
        print(f"  ... 他 {len(pending) - 6} 件")

    next_topic = pending[0]
    print(f"\n次のトピック → ID {next_topic['id']}: {next_topic['theme']}")
    choice = input("Enterでこのまま生成 / 別のIDを指定する場合は番号を入力: ").strip()

    if choice.isdigit():
        matched = [t for t in data["topics"] if t["id"] == int(choice)]
        if matched:
            next_topic = matched[0]

    research_text = generate_script.load_latest_research()
    print("\nClaude APIに接続中...")
    try:
        script_text = generate_script.generate_script(next_topic, research_text)
        filepath = generate_script.save_script(next_topic, script_text)
    except Exception as e:
        print(f"\nエラー: {e}")
        return

    for t in data["topics"]:
        if t["id"] == next_topic["id"]:
            t["done"] = True
    generate_script.save_topics(data)
    print(f"\n台本を保存しました: {filepath.relative_to(BASE_DIR)}")


# -----------------------------------------------------------------
# メニュー5: スピーキングショート動画台本生成（APIキー必要）
# -----------------------------------------------------------------

def run_generate_speaking_short() -> None:
    if not ensure_api_key():
        print("APIキーが設定されていないため、この機能は使えません。")
        return

    try:
        import generate_speaking_short
    except ImportError as e:
        print(f"エラー: {e}")
        return

    print()
    print("【スピーキングショート動画台本生成】")
    print("採点官・受験生・コーチの3視点で構成されるショート動画台本を生成します。")
    print()

    data = generate_speaking_short.load_topics()
    pending = [t for t in data["topics"] if not t["done"]]

    if not pending:
        print("すべてのスピーキングトピックが完了しています。")
        reset = input("リセットしますか？ (y/n): ").strip().lower()
        if reset == "y":
            for t in data["topics"]:
                t["done"] = False
            generate_speaking_short.save_topics(data)
            pending = data["topics"]
        else:
            return

    print("未完了のスピーキングトピック:")
    for t in pending[:6]:
        print(f"  [{t['id']:2d}] Part {t['part']} | {t['theme']:<20} | {t['question'][:45]}...")
    if len(pending) > 6:
        print(f"  ... 他 {len(pending) - 6} 件")

    print()
    print("c. カスタム質問を入力する")
    choice = input("Enterで次のトピック / IDを指定 / c でカスタム: ").strip()

    if choice.lower() == "c":
        question = input("スピーキングの質問を入力してください: ").strip()
        if not question:
            print("キャンセルしました。")
            return
        part_str = input("Part番号 (1/2/3、デフォルト1): ").strip()
        part = int(part_str) if part_str.isdigit() else 1
        grammar = input("文法フォーカス（空でデフォルト）: ").strip()
        topic = {
            "id": "custom",
            "part": part,
            "theme": "Custom",
            "question": question,
            "grammar_focus": grammar or "Mixed Conditional / Advanced vocabulary",
            "done": False,
        }
    else:
        if choice.isdigit():
            matched = [t for t in data["topics"] if t["id"] == int(choice)]
            topic = matched[0] if matched else pending[0]
        else:
            topic = pending[0]

    print("\nClaude APIに接続中...")
    try:
        script_text = generate_speaking_short.generate_script(topic)
        filepath = generate_speaking_short.save_script(topic, script_text)
    except Exception as e:
        print(f"\nエラー: {e}")
        return

    if topic["id"] != "custom":
        for t in data["topics"]:
            if t["id"] == topic["id"]:
                t["done"] = True
        generate_speaking_short.save_topics(data)

    print(f"\n台本を保存しました: {filepath.relative_to(BASE_DIR)}")


# -----------------------------------------------------------------
# メニュー6: トピック一覧
# -----------------------------------------------------------------

def show_topics_list() -> None:
    if not TOPICS_FILE.exists():
        print("エラー: topics.json が見つかりません。")
        return

    with open(TOPICS_FILE, encoding="utf-8") as f:
        data = json.load(f)

    topics = data["topics"]
    done_count = sum(1 for t in topics if t["done"])

    print()
    print(f"トピック一覧  （残り {len(topics) - done_count} 本 / 全 {len(topics)} 本）")
    print()
    for t in topics:
        mark = "✓" if t["done"] else "○"
        tag = ""
        if t.get("source") == "seo_research":
            tag = " [SEO]"
        elif t.get("source") == "correction_video":
            tag = " [添削]"
        print(f"  [{mark}] ID {t['id']:2d} | {t['theme']:<30} | {t['essay_type']}{tag}")
    print()


# -----------------------------------------------------------------
# エントリーポイント
# -----------------------------------------------------------------

def main() -> None:
    load_env_file()

    while True:
        choice = show_main_menu()
        if choice == "1":
            run_correction_workflow()
        elif choice == "2":
            save_seo_research()
        elif choice == "3":
            run_seo_research_auto()
        elif choice == "4":
            run_generate_script_auto()
        elif choice == "5":
            run_generate_speaking_short()
        elif choice == "6":
            show_topics_list()
        elif choice == "7":
            print("\n終了します。\n")
            break


if __name__ == "__main__":
    main()
