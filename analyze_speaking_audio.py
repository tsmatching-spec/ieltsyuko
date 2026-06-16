#!/usr/bin/env python3
"""
IELTS Speaking Audio → Short Video Script Generator

Takes an audio/video recording of a coaching session (student answer + coach feedback)
and generates a short video script in 採点官・受験生・コーチ format.

Usage:
    python analyze_speaking_audio.py --audio path/to/file.mp4
    python analyze_speaking_audio.py --audio path/to/file.mp3 --part 2
    python analyze_speaking_audio.py --transcript  # paste transcript manually
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

SCRIPTS_DIR = Path(__file__).parent / "scripts" / "speaking_shorts"
SUPPORTED_AUDIO = {".mp3", ".mp4", ".m4a", ".wav", ".webm", ".ogg"}


# -----------------------------------------------------------------
# 文字起こし
# -----------------------------------------------------------------

def transcribe_with_whisper(audio_path: Path) -> str:
    """Transcribe audio using OpenAI Whisper API."""
    try:
        import openai
    except ImportError:
        print("エラー: openai パッケージが見つかりません。")
        print("  pip install openai  を実行してください。")
        sys.exit(1)

    openai_key = os.environ.get("OPENAI_API_KEY", "")
    if not openai_key:
        env_file = Path(__file__).parent / ".env"
        print("\n  OpenAI APIキー（Whisper用）が必要です。")
        print("  取得: https://platform.openai.com/api-keys")
        key = input("  OpenAI APIキーを貼り付けてください: ").strip()
        if not key:
            print("キャンセルしました。")
            sys.exit(1)
        with open(env_file, "a", encoding="utf-8") as f:
            f.write(f"\nOPENAI_API_KEY={key}\n")
        os.environ["OPENAI_API_KEY"] = key
        openai_key = key

    client = openai.OpenAI(api_key=openai_key)

    print(f"文字起こし中: {audio_path.name} ...")
    with open(audio_path, "rb") as f:
        response = client.audio.transcriptions.create(
            model="whisper-1",
            file=f,
            language="ja",
            response_format="text"
        )
    return response


def get_transcript_from_paste() -> str:
    """Manually paste transcript."""
    print()
    print("文字起こしテキストを貼り付けてください。")
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
    return "\n".join(lines).strip()


# -----------------------------------------------------------------
# Claude による分析 → 台本生成
# -----------------------------------------------------------------

def build_analysis_prompt(transcript: str, part: int) -> str:
    return f"""あなたはIELTSスピーキング指導のプロで、YouTubeチャンネル「IELTSゆうこ」のショート動画台本ライターです。

以下は、IELTSスピーキングのコーチング音声の文字起こしです。
コーチ（ゆうこ）が生徒の回答にフィードバックしている内容が含まれています。

---
{transcript}
---

この文字起こしから以下の情報を抽出し、ショート動画台本を作成してください。

**抽出する情報：**
1. IELTSスピーキングの質問（何について答えているか）
2. 生徒の回答（文字起こしから読み取れる範囲で）
3. コーチが指摘した問題点・弱点
4. コーチが提案した改善点・テクニック
5. 文法フォーカス（コーチが強調した文法・語彙）

**出力フォーマット（Markdown形式）**

# [動画タイトル（日本語、IELTS学習者の興味を引くタイトル）]

## 📋 動画概要
- **対象**: IELTS Speaking Part {part}
- **テーマ**: （テーマ名）
- **文法フォーカス**: （コーチが強調した文法・語彙）
- **想定時間**: 45〜60秒

---

## 🎬 台本

### オープニング（0〜5秒）

**【採点官】**:（採点官として視聴者に語りかける一言。今回のポイントを一言で）

**【コーチ】**:（今日のテーマと学ぶポイントを一言で紹介）

---

### 本編（5〜50秒）

**【コーチ】**: 今日の質問はこちら！

> 💬 質問テロップ: "（抽出した質問文）"

**【受験生（初級）】**:（文字起こしから読み取った生徒の実際の答えをベースにした初級レベルの答え）

**【採点官（心の声）】**:（採点官として、この答えに感じる物足りなさをひとことで）

**【コーチ】**:（コーチが実際に指摘した改善ポイントをベースに一言。「〇〇を使うと自然に差がつきます！」スタイル）

> 📌 テクニックテロップ: （コーチが教えたテクニックの名前・キーワード）

**【受験生（中級）】**:（コーチのアドバイスを活かした中級レベルの答え）

**【コーチ】**:（さらに上のレベルへのヒント。コーチが実際に提案した上級テクニック）

> 📌 テクニックテロップ: （上級テクニックの名前・キーワード）

**【受験生（上級）】**:（コーチの最高レベルのアドバイスを活かした上級の答え）

---

### エンディング（50〜60秒）

**【採点官】**:（上級の答えへの評価。「Band 7以上ねらえます！」など具体的なスコア感）

**【コーチ】**:（視聴者へのひとこと締め。「試してみて！」「保存して練習しよう！」など）

---

## 📝 スクリプト解説

### 初級 → 中級 → 上級の変化ポイント

| レベル | 答えの例 | 使っているテクニック |
|--------|---------|---------------------|
| 初級 | （初級の答えを再掲） | シンプルな表現のみ |
| 中級 | （中級の答えを再掲） | （使ったテクニック） |
| 上級 | （上級の答えを再掲） | （コーチが強調した文法・語彙） |

### コーチのワンポイント解説
（コーチが実際に指摘した内容を日本語で100字以内でまとめる）

### 採点官が見ているポイント（Part {part}）
（Part {part}で採点官が特に重視する3つのポイントを箇条書きで）

---

## 🎨 撮影・編集メモ
- **画面構成**: 採点官（右上）・受験生（中央）・コーチ（左下）の3分割
- **テロップスタイル**: 質問は大きく白抜き、テクニック名はオレンジ強調
- **BGM**: 軽快でテンポよい学習系BGM
- **字幕**: 全セリフに日本語字幕（英語発話部分は英語+日本語訳）
- **エンド画面**: 「保存して練習！」CTAと関連動画サムネ
"""


def generate_script_from_transcript(transcript: str, part: int) -> str:
    client = anthropic.Anthropic()

    print("Claude APIで台本を生成中...")

    message = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        messages=[
            {"role": "user", "content": build_analysis_prompt(transcript, part)}
        ]
    )

    return message.content[0].text


def save_script(script_text: str, source_name: str, part: int) -> Path:
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = Path(source_name).stem[:30].replace(" ", "_")
    filename = f"{date_str}_from_audio_part{part}_{stem}.md"
    filepath = SCRIPTS_DIR / filename

    header = f"""---
generated: {datetime.now().isoformat()}
type: speaking_short_from_audio
source_audio: {source_name}
part: {part}
---

"""
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(header + script_text)

    return filepath


# -----------------------------------------------------------------
# エントリーポイント
# -----------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="IELTS Speaking コーチング音源 → ショート動画台本生成"
    )
    parser.add_argument("--audio", metavar="FILE", help="音声・動画ファイルのパス (mp3/mp4/m4a/wav)")
    parser.add_argument("--transcript", action="store_true", help="文字起こしを手動で貼り付ける")
    parser.add_argument("--part", type=int, default=1, choices=[1, 2, 3],
                        help="IELTS Speaking Part番号 (デフォルト: 1)")
    args = parser.parse_args()

    if args.audio:
        audio_path = Path(args.audio)
        if not audio_path.exists():
            print(f"エラー: ファイルが見つかりません: {args.audio}")
            sys.exit(1)
        if audio_path.suffix.lower() not in SUPPORTED_AUDIO:
            print(f"エラー: 対応していないファイル形式です: {audio_path.suffix}")
            print(f"対応形式: {', '.join(SUPPORTED_AUDIO)}")
            sys.exit(1)
        transcript = transcribe_with_whisper(audio_path)
        source_name = audio_path.name
        print(f"\n--- 文字起こし結果 ---\n{transcript[:300]}...\n")

    elif args.transcript:
        transcript = get_transcript_from_paste()
        source_name = "manual_transcript"
        if not transcript:
            print("テキストが入力されませんでした。")
            sys.exit(1)

    else:
        # インタラクティブモード
        print()
        print("【コーチング音源 → ショート動画台本生成】")
        print()
        print("1. 音声ファイルを指定する（Whisper文字起こし）")
        print("2. 文字起こしテキストを貼り付ける")
        print()
        mode = input("選択 (1/2): ").strip()

        if mode == "1":
            audio_input = input("音声ファイルのパス: ").strip().strip('"')
            audio_path = Path(audio_input)
            if not audio_path.exists():
                print(f"エラー: ファイルが見つかりません: {audio_input}")
                sys.exit(1)
            transcript = transcribe_with_whisper(audio_path)
            source_name = audio_path.name
            print(f"\n--- 文字起こし結果（先頭300文字）---\n{transcript[:300]}...\n")
        else:
            transcript = get_transcript_from_paste()
            source_name = "manual_transcript"
            if not transcript:
                print("テキストが入力されませんでした。")
                sys.exit(1)

        part_input = input("Part番号 (1/2/3、デフォルト1): ").strip()
        args.part = int(part_input) if part_input in ("1", "2", "3") else 1

    script_text = generate_script_from_transcript(transcript, args.part)
    filepath = save_script(script_text, source_name, args.part)

    print(f"\n台本を保存しました: {filepath.relative_to(Path(__file__).parent)}")
    print("Done!")


if __name__ == "__main__":
    main()
