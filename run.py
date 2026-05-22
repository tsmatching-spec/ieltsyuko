#!/usr/bin/env python3
"""
IELTSゆうこ 動画制作ツール
起動: python run.py
"""

import os
import sys
import json
from pathlib import Path

ENV_FILE = Path(__file__).parent / ".env"
TOPICS_FILE = Path(__file__).parent / "topics.json"


# -----------------------------------------------------------------
# .env ローダー（モジュールインポート前に実行する必要がある）
# -----------------------------------------------------------------

def load_env_file() -> None:
    if not ENV_FILE.exists():
        return
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def check_or_setup_api_key() -> bool:
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key and not key.startswith("sk-ant-xxx"):
        return True

    print()
    print("=" * 50)
    print("  初回セットアップ: Anthropic APIキーの設定")
    print("=" * 50)
    print()
    print("このツールを使うには Anthropic の APIキーが必要です。")
    print()
    print("取得方法:")
    print("  1. https://console.anthropic.com/ にアクセス")
    print("  2. ログイン後、「API Keys」からキーを発行")
    print("  3. 「sk-ant-」で始まる文字列をコピー")
    print()

    while True:
        key = input("APIキーを貼り付けてください: ").strip()
        if not key:
            print("  APIキーを入力してください。")
            continue
        if not key.startswith("sk-ant-"):
            print("  ※ APIキーは「sk-ant-」で始まります。もう一度確認してください。")
            cont = input("  このまま続けますか？ (y/n): ").strip().lower()
            if cont != "y":
                continue
        break

    # .env に書き込む
    env_content = f"ANTHROPIC_API_KEY={key}\n"
    ENV_FILE.write_text(env_content, encoding="utf-8")
    os.environ["ANTHROPIC_API_KEY"] = key
    print()
    print("  APIキーを .env に保存しました。次回から自動で読み込まれます。")
    return True


# -----------------------------------------------------------------
# メニュー
# -----------------------------------------------------------------

def show_main_menu() -> str:
    print()
    print("=" * 40)
    print("  IELTSゆうこ 動画制作ツール")
    print("=" * 40)
    print()
    print("  1. SEOリサーチを実行する（月1回）")
    print("  2. 台本を1本生成する")
    print("  3. 残りのトピック一覧を見る")
    print("  4. 終了")
    print()
    while True:
        choice = input("番号を入力してください (1〜4): ").strip()
        if choice in ("1", "2", "3", "4"):
            return choice
        print("  有効な番号を入力してください（1〜4）")


# -----------------------------------------------------------------
# メニュー1: SEOリサーチ
# -----------------------------------------------------------------

def run_seo_research() -> None:
    import research_seo

    print()
    print("【SEOリサーチ】")
    print("YouTubeでIELTSライティングの検索需要・競合を分析し、")
    print("次に作るべき動画の企画をレポートにまとめます。")
    print()

    add = input("完了後、おすすめ企画を topics.json に自動追加しますか？ (y/n): ").strip().lower()
    add_topics = add == "y"

    import anthropic
    try:
        client = anthropic.Anthropic()
    except Exception as e:
        print(f"\nエラー: APIクライアントの初期化に失敗しました。\n{e}")
        return

    queries = research_seo.SEARCH_QUERIES + research_seo.COMPETITOR_QUERIES
    print(f"\n検索を開始します（{len(queries)} 件）...\n")

    try:
        search_results = research_seo.run_web_searches(client, queries)
    except Exception as e:
        print(f"\nエラー: 検索中に問題が発生しました。\n{e}")
        return

    print("\nレポートを生成中...")
    try:
        report_text = research_seo.synthesize_report(client, search_results)
        report_path = research_seo.save_report(report_text)
    except Exception as e:
        print(f"\nエラー: レポート生成に失敗しました。\n{e}")
        return

    print(f"\nレポートを保存しました: {report_path.relative_to(Path(__file__).parent)}")

    if add_topics:
        print("\n企画を topics.json に追加中...")
        try:
            existing_ids = research_seo.get_existing_ids()
            new_topics = research_seo.extract_topics_from_report(client, report_text, existing_ids)
            added = research_seo.add_topics_to_json(new_topics)
            print(f"{added} 件の企画を topics.json に追加しました。")
        except Exception as e:
            print(f"※ 自動追加に失敗しました（手動でレポートを確認してください）: {e}")

    print("\n次のステップ: メニュー「2」で台本を生成してください。")


# -----------------------------------------------------------------
# メニュー2: 台本生成
# -----------------------------------------------------------------

def run_generate_script() -> None:
    import generate_script

    print()
    print("【台本生成】")

    # 未完了トピックを取得
    try:
        data = generate_script.load_topics()
    except FileNotFoundError:
        print("エラー: topics.json が見つかりません。")
        return

    pending = [t for t in data["topics"] if not t["done"]]

    if not pending:
        print("すべてのトピックが完了しています。")
        print("メニュー「1」でSEOリサーチを実行して新しい企画を追加するか、")
        done_reset = input("リセットして最初からやり直しますか？ (y/n): ").strip().lower()
        if done_reset == "y":
            for t in data["topics"]:
                t["done"] = False
            generate_script.save_topics(data)
            print("リセットしました。もう一度「2」を選んでください。")
        return

    # トピック一覧表示
    print()
    print("未完了のトピック:")
    for t in pending[:6]:
        print(f"  [{t['id']:2d}] {t['theme']:<25} {t['essay_type']}")
    if len(pending) > 6:
        print(f"  ... 他 {len(pending) - 6} 件")

    # トピック選択
    next_topic = pending[0]
    print()
    print(f"次のトピック → ID {next_topic['id']}: {next_topic['theme']} ({next_topic['essay_type']})")
    choice = input("Enterでこのまま生成 / 別のIDを指定する場合は番号を入力: ").strip()

    if choice:
        if choice.isdigit():
            topic_id = int(choice)
            matched = [t for t in data["topics"] if t["id"] == topic_id]
            if matched:
                next_topic = matched[0]
            else:
                print(f"  ID {topic_id} は見つかりません。最初のトピックを使います。")
        else:
            print("  数字以外が入力されました。最初のトピックを使います。")

    print(f"\nトピック「{next_topic['theme']}」で台本を生成します...")

    # SEOコンテキスト読み込み
    research_text = generate_script.load_latest_research()
    if research_text:
        print("  SEOリサーチ情報: 読み込み済み")
    else:
        print("  SEOリサーチ情報: なし（先にメニュー「1」を実行するとSEO最適化されます）")

    print("\nClaude APIに接続中...")
    try:
        script_text = generate_script.generate_script(next_topic, research_text)
        filepath = generate_script.save_script(next_topic, script_text)
    except Exception as e:
        print(f"\nエラー: 台本生成に失敗しました。\n{e}")
        return

    # 完了フラグを保存
    for t in data["topics"]:
        if t["id"] == next_topic["id"]:
            t["done"] = True
    generate_script.save_topics(data)

    rel_path = filepath.relative_to(Path(__file__).parent)
    print(f"\n台本を保存しました: {rel_path}")
    print("テキストエディタやメモ帳で開いて内容を確認してください。")


# -----------------------------------------------------------------
# メニュー3: トピック一覧
# -----------------------------------------------------------------

def show_topics_list() -> None:
    import generate_script

    try:
        data = generate_script.load_topics()
    except FileNotFoundError:
        print("エラー: topics.json が見つかりません。")
        return

    topics = data["topics"]
    done_count = sum(1 for t in topics if t["done"])

    print()
    print(f"トピック一覧  （残り {len(topics) - done_count} 本 / 全 {len(topics)} 本）")
    print()
    for t in topics:
        mark = "✓" if t["done"] else "○"
        src = " [SEO]" if t.get("source") == "seo_research" else ""
        print(f"  [{mark}] ID {t['id']:2d} | {t['theme']:<25} | {t['essay_type']}{src}")
    print()


# -----------------------------------------------------------------
# エントリーポイント
# -----------------------------------------------------------------

def main() -> None:
    # .env を読み込んでからモジュールをインポートする
    load_env_file()

    try:
        import anthropic  # noqa: F401
    except ImportError:
        print("エラー: anthropic ライブラリがインストールされていません。")
        print("以下のコマンドを実行してください:")
        print("  pip install anthropic")
        sys.exit(1)

    if not check_or_setup_api_key():
        sys.exit(1)

    while True:
        choice = show_main_menu()
        if choice == "1":
            run_seo_research()
        elif choice == "2":
            run_generate_script()
        elif choice == "3":
            show_topics_list()
        elif choice == "4":
            print("\n終了します。\n")
            break


if __name__ == "__main__":
    main()
