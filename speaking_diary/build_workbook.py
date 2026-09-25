#!/usr/bin/env python3
"""
IELTS スピーキング日記 スプレッドシート生成スクリプト
起動: python speaking_diary/build_workbook.py
出力: speaking_diary/IELTS_Speaking_Diary.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).parent / "IELTS_Speaking_Diary.xlsx"

FONT = "Arial"
NAVY = "1F3864"
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
HEADER_FILL = PatternFill("solid", fgColor="EDEDED")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # 自分で入力するセル
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

WPM = 140  # 2分スピーチの目安話速（words per minute）


def f(bold=False, size=10, color="000000"):
    return Font(name=FONT, bold=bold, size=size, color=color)


def word_count(ref):
    return (f'=IF(LEN(TRIM({ref}))=0,0,'
            f'LEN(TRIM({ref}))-LEN(SUBSTITUTE(TRIM({ref})," ",""))+1)')


# ---------------------------------------------------------------
# 例（2026-09-25）の中身
# ---------------------------------------------------------------
EXAMPLE = {
    "date": "2026-09-25",
    "jp": ("今日は朝から雨で、駅に着いたら人身事故で電車が40分遅れていた。"
           "最初はイライラしたけど、駅前のカフェに入ってIELTSの単語を復習した。"
           "意外と集中できて、充実した時間になった。会社に遅れると連絡したら、"
           "上司が「今日はリモートでいいよ」と言ってくれて、家に戻って仕事をした。"),
    "feel": ("最初はイライラ→結果的にラッキーだった。"
             "毎日の通勤って本当に必要なのかな、と考えさせられた。"),
    "b1": ("It was raining this morning. When I got to the station, the trains were "
           "40 minutes late because of an accident. At first I was annoyed, but I went "
           "into a café near the station and studied my IELTS vocabulary. I could "
           "concentrate very well, so it was a good time. When I told my company that I "
           "would be late, my boss said, \"You can work from home today.\" So I went back "
           "home and worked there. It made me think about whether we really need to go "
           "to the office every day."),
    "b2": ("It had been pouring since early morning, and by the time I reached the "
           "station, the trains were running forty minutes behind schedule due to an "
           "accident. Although I was frustrated at first, I decided to make the most of "
           "the delay by reviewing IELTS vocabulary in a café near the station. "
           "Surprisingly, I was able to focus far better than I usually do at home. When I "
           "contacted my office to say I'd be late, my manager kindly suggested that I "
           "work remotely instead, so I headed back home. The whole experience made me "
           "question whether commuting every day is really necessary."),
    "vocab": [
        ("make the most of ~", "〜を最大限に活用する", "慣用表現 → Lexical Resource で高評価",
         "I decided to make the most of the delay."),
        ("a blessing in disguise", "災い転じて福となす", "イディオムを自然に使える印象",
         "The delay turned out to be a blessing in disguise."),
        ("behind schedule", "予定より遅れて", "late の言い換え（語彙の幅）",
         "The trains were running behind schedule."),
        ("work remotely / hybrid working", "リモート勤務／ハイブリッド勤務", "Part 3 の社会テーマに直結",
         "More companies now allow hybrid working."),
        ("commute (n./v.)", "通勤（する）", "Part 1 Work/Transport の必須語",
         "My daily commute takes about an hour."),
        ("It made me question whether ~", "〜かどうか考えさせられた", "意見・内省を示す定型",
         "It made me question whether commuting is necessary."),
        ("far better than ~", "〜よりはるかに良い", "比較の強調（far / much）",
         "I could focus far better than I do at home."),
        ("counterproductive", "逆効果の", "Band 7+ 語彙。Part 3 の議論で有効",
         "Forcing everyone into the office can be counterproductive."),
    ],
    "themes": [
        ("リモートワーク／ハイブリッド勤務", "https://news.google.com/search?q=hybrid%20work%20return%20to%20office&hl=en",
         "コロナ禍以降、多くの企業がハイブリッド勤務を導入。一方で近年は協働や社風維持を理由に"
         "出社回帰を求める企業も増え、「生産性」と「柔軟性」のどちらを重視するかが議論になっている。"),
        ("公共交通の信頼性・通勤", "https://news.google.com/search?q=Japan%20train%20punctuality%20commute&hl=en",
         "日本の鉄道は定時運行で世界的に知られ、遅延そのものがニュースになるほど。"
         "首都圏では片道1時間前後の通勤も珍しくなく、混雑・ストレス・時間の損失が社会課題として語られる。"),
        ("時間の有効活用・スキマ時間学習", "https://news.google.com/search?q=microlearning%20productivity&hl=en",
         "短時間の学習（マイクロラーニング）は集中力を保ちやすく、記憶の定着にも効果的とされる。"
         "待ち時間をどう使うかは Part 2/3 の「時間の使い方」系トピックで使える。"),
    ],
    "themes_note": ("※URLはニュース検索リンクです。実際に読んだ記事のURLと要約に差し替えてください"
                    "（要約は一般的な背景知識のまとめです）。"),
    "placement": [
        ("Part 1", "Do you work or study? / How do you usually travel to work? / Do you like rainy days?",
         "「今朝まさに電車が遅れて…」と具体例を1文添えて回答を2〜3文に広げる。"),
        ("Part 2", "Describe a time when you had to wait for something. / "
                   "Describe an unexpected event that turned out well.",
         "今日の話をそのまま使える。雨→遅延→カフェ→リモートの流れで時系列に話す。"
         "最後に「何を学んだか」で締める。"),
        ("Part 3", "Is working from home better than working in an office? / "
                   "How can governments improve public transport?",
         "個人の体験 → 社会一般の議論へ広げる。メリット・デメリットを両方述べて自分の立場を示す。"),
    ],
    "grammar": [
        ("時制の幅",
         "It had been pouring since early morning. / The trains were running late. / "
         "I've started to wonder whether... / I'll probably ask my manager about it.",
         "過去完了進行・過去進行・現在完了・未来を1つの話に混ぜる（Grammatical Range）"),
        ("仮定法",
         "If the train hadn't been delayed, I would never have discovered how productive a café can be. / "
         "If I could work from home twice a week, I would study more.",
         "仮定法過去完了（過去の反実仮想）＋仮定法過去（現在の仮定）をセットで"),
        ("助動詞",
         "I should have left home earlier. / It might have been the best part of my day. / "
         "Companies should offer more flexible options.",
         "助動詞＋have p.p. で後悔・推量、should で提案・意見"),
        ("複雑な文構造",
         "What surprised me most was how well I could concentrate. / "
         "Having nothing else to do, I reviewed fifty words. / "
         "Not only does commuting take time, but it is also stressful.",
         "強調構文(What ~ is)・分詞構文・倒置・関係詞(which turned out to be...)"),
    ],
    "speech": (
        "I'd like to talk about something that happened to me this morning, which turned out "
        "to be a blessing in disguise. It had been pouring since early morning, and by the time "
        "I reached the station, the trains were running about forty minutes behind schedule "
        "because of an accident. To be honest, my first reaction was frustration, and I kept "
        "thinking that I should have left home earlier. However, instead of standing on a "
        "crowded platform, I decided to make the most of the delay. I went into a small café "
        "near the station, ordered a coffee, and reviewed my IELTS vocabulary. What surprised "
        "me most was how well I could concentrate. Having nothing else to do, I managed to go "
        "through about fifty new words, which is far more than I usually learn in an evening "
        "at home. Then, when I contacted my office to say I would be late, my manager kindly "
        "suggested that I work remotely for the rest of the day, so I headed back home and "
        "finished my tasks there. Looking back, if the train hadn't been delayed, I would "
        "never have discovered how productive a short break in a café can be. It also made me "
        "question whether commuting every single day is really necessary. Not only does it "
        "take up a lot of time, but it can also be quite stressful, especially in bad weather. "
        "I think companies should offer more flexible options, such as hybrid working, so that "
        "employees can choose the environment in which they work best. All in all, what "
        "started as a frustrating morning ended up being one of the most productive days I've "
        "had recently."),
}


# ---------------------------------------------------------------
# 1日分のページ（縦型レイアウト）
# ---------------------------------------------------------------
def build_day_sheet(ws, d=None):
    """d=None なら空テンプレート。戻り値: スピーチ本文セル座標"""
    d = d or {}
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", (18, 30, 30, 30, 40)):
        ws.column_dimensions[col].width = w

    r = 1
    ws.cell(r, 1, "IELTS Speaking Diary｜今日の出来事 → 英語スピーチ").font = f(True, 14, NAVY)
    r += 2

    def section(title, note=""):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, f"{title}　{note}" if note else title)
        c.font = f(True, 11, NAVY)
        c.fill = SECTION_FILL
        ws.row_dimensions[r].height = 20
        r += 1

    def long_row(label, value, height, is_input=False):
        nonlocal r
        ws.cell(r, 1, label).font = f(True)
        ws.cell(r, 1).alignment = WRAP
        ws.cell(r, 1).border = BOX
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(r, 2, value)
        c.font = f()
        c.alignment = WRAP
        if is_input:
            c.fill = INPUT_FILL
        for col in range(2, 6):
            ws.cell(r, col).border = BOX
        ws.row_dimensions[r].height = height
        r += 1
        return f"B{r - 1}"

    def table(headers, rows, n_blank, height=45, merge_last=False):
        """headers: A〜D(E) の見出し。merge_last=True なら最後の列を D:E に結合"""
        nonlocal r
        cols = len(headers)
        for i, h in enumerate(headers, 1):
            c = ws.cell(r, i, h)
            c.font = f(True)
            c.fill = HEADER_FILL
            c.alignment = CENTER
            c.border = BOX
        if merge_last:
            ws.merge_cells(start_row=r, start_column=cols, end_row=r, end_column=5)
            ws.cell(r, 5).border = BOX
        r += 1
        data = list(rows) + [("",) * cols] * max(0, n_blank - len(rows))
        for row in data:
            for i, v in enumerate(row, 1):
                c = ws.cell(r, i, v)
                c.font = f(True) if i == 1 else f()
                c.alignment = WRAP
                c.border = BOX
            if merge_last:
                ws.merge_cells(start_row=r, start_column=cols, end_row=r, end_column=5)
                ws.cell(r, 5).border = BOX
            ws.row_dimensions[r].height = height
            r += 1

    # 0. 基本情報
    section("① 今日の日記（日本語で入力）", "← 黄色セルに入力")
    long_row("日付", d.get("date", ""), 18, is_input=True)
    long_row("今日あったこと", d.get("jp", ""), 90, is_input=True)
    long_row("感じたこと・考えたこと", d.get("feel", ""), 45, is_input=True)
    r += 1

    # 1. 英訳
    section("② 英語訳", "B1 = シンプルで正確 / B2 = 接続詞・言い換えで幅を出す")
    long_row("B1 レベル", d.get("b1", ""), 110)
    long_row("B2 レベル", d.get("b2", ""), 120)
    r += 1

    # 2. アピール語彙
    section("③ 使うべき単語・アピールすべき表現")
    table(("単語・表現", "意味", "なぜアピールになるか", "今日の内容での例文"),
          d.get("vocab", []), 8, height=32, merge_last=True)
    r += 1

    # 3. 関連テーマ・ニュース
    section("④ 関連テーマ・ニュース記事と要約")
    table(("関連テーマ", "URL（ニュース記事）", "記事の要約"),
          d.get("themes", []), 3, height=70, merge_last=True)
    if d.get("themes_note"):
        ws.cell(r, 1, d["themes_note"]).font = f(size=9, color="7F7F7F")
        r += 1
    r += 1

    # 4. スピーキングでの使い所
    section("⑤ IELTS スピーキングのどこで使うか")
    table(("Part", "想定される質問", "この内容の入れ方"),
          d.get("placement", []), 3, height=60, merge_last=True)
    r += 1

    # 5. 文法アピール
    section("⑥ 文法アピールポイント", "時制の幅・仮定法・助動詞・複雑な文構造")
    table(("文法項目", "今日の内容で使える例文", "ポイント"),
          d.get("grammar", []), 4, height=62, merge_last=True)
    r += 1

    # 6. 2分スピーチ
    section("⑦ 2分スピーチ 完成版（英語）")
    speech_ref = long_row("スピーチ本文", d.get("speech", ""), 330)
    ws.cell(r, 1, "語数").font = f(True)
    ws.cell(r, 2, word_count(speech_ref)).font = f()
    ws.cell(r, 3, f"目安時間（{WPM} wpm）").font = f(True)
    wc_ref = f"B{r}"
    c = ws.cell(r, 4, f"=IF({wc_ref}=0,\"\",TEXT({wc_ref}/{WPM}/1440,\"m:ss\"))")
    c.font = f()
    ws.cell(r, 5, f"2分 ≒ {2 * WPM} 語が目安").font = f(size=9, color="7F7F7F")
    r += 1
    ws.cell(r, 1, "本番練習の録音時間").font = f(True)
    ws.cell(r, 2).fill = INPUT_FILL
    ws.cell(r, 3, "自己評価（1〜5）").font = f(True)
    ws.cell(r, 4).fill = INPUT_FILL
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5",
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(ws.cell(r, 4))
    ws.freeze_panes = "A3"
    return speech_ref


# ---------------------------------------------------------------
# 使い方
# ---------------------------------------------------------------
def build_guide(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 100
    ws["B1"] = "IELTS Speaking Diary の使い方"
    ws["B1"].font = f(True, 14, NAVY)
    lines = [
        ("1", "「テンプレート」シートを右クリック →「コピーを作成」し、シート名を日付（例: 2026-09-26）に変更。"),
        ("2", "① の黄色セルに、今日あったこと・感じたことを日本語で入力。"),
        ("3", "② 〜 ⑦ を埋める（Claude に①の日本語を貼って「この形式で埋めて」と頼むと早いです）。"),
        ("4", "⑦ のスピーチを声に出して練習し、録音時間と自己評価を記入。語数・目安時間は自動計算。"),
        ("5", "「記録一覧」に1行追加し、よく使った表現は「単語帳」に転記して繰り返し使う。"),
        ("", ""),
        ("凡例", "黄色セル = 自分で入力するセル ／ 青い帯 = セクション見出し ／ 語数・目安時間 = 自動計算"),
        ("見本", "「2026-09-25（例）」シートに記入例があります。"),
    ]
    for i, (n, t) in enumerate(lines, 3):
        ws.cell(i, 1, n).font = f(True)
        ws.cell(i, 2, t).font = f()
        ws.cell(i, 2).alignment = WRAP
    ws["A10"].fill = INPUT_FILL


# ---------------------------------------------------------------
# 記録一覧・単語帳
# ---------------------------------------------------------------
def header_row(ws, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, i, h)
        c.font = f(True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = CENTER
        c.border = BOX
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def build_log(ws, example_sheet, speech_ref):
    header_row(ws, ("日付", "シート名", "トピック（一言）", "使えるPart", "語数",
                    "録音時間", "自己評価(1-5)", "メモ"),
               (12, 18, 30, 12, 8, 10, 12, 40))
    q = f"'{example_sheet}'"
    ex = ["2026-09-25", example_sheet, "電車遅延→カフェ学習→リモート勤務", "Part 2 / 3",
          f"={q}!B{int(speech_ref[1:]) + 1}", "", "", "例：仮定法過去完了を自然に言えた"]
    for i, v in enumerate(ex, 1):
        c = ws.cell(2, i, v)
        c.font = f()
        c.border = BOX
        c.alignment = WRAP
    for row in range(3, 32):
        for col in range(1, 9):
            ws.cell(row, col).border = BOX
            ws.cell(row, col).font = f()


def build_vocab(ws, vocab):
    header_row(ws, ("日付", "単語・表現", "意味", "アピールポイント", "例文", "使用回数"),
               (12, 28, 22, 32, 50, 10))
    for i, (w, m, why, ex) in enumerate(vocab, 2):
        for j, v in enumerate(("2026-09-25", w, m, why, ex, ""), 1):
            c = ws.cell(i, j, v)
            c.font = f()
            c.border = BOX
            c.alignment = WRAP


def build_grammar_ref(ws):
    header_row(ws, ("カテゴリ", "文法項目", "型", "例文", "IELTSでの効果"),
               (14, 22, 30, 55, 34))
    rows = [
        ("時制の幅", "過去完了", "had + p.p.", "By the time I arrived, the train had already left.", "出来事の前後関係を明確にできる"),
        ("時制の幅", "過去完了進行", "had been + -ing", "It had been raining for hours.", "継続の背景描写（Part 2 の導入に最適）"),
        ("時制の幅", "現在完了（進行）", "have (been) + p.p./-ing", "I've been studying for IELTS since April.", "経験・継続を自然に表す"),
        ("時制の幅", "未来進行・未来完了", "will be -ing / will have p.p.", "By next year, I will have finished my degree.", "Band 7+ でも差がつく時制"),
        ("仮定法", "仮定法過去", "If + 過去形, would + 原形", "If I lived closer to work, I would walk.", "現在の仮定・理想（Part 3 の意見）"),
        ("仮定法", "仮定法過去完了", "If + had p.p., would have p.p.", "If I hadn't missed the train, I wouldn't have met her.", "過去の振り返り（Part 2 の締め）"),
        ("仮定法", "混合仮定法", "If + had p.p., would + 原形", "If I had studied abroad, I would be more confident now.", "上級者アピール"),
        ("仮定法", "I wish / If only", "I wish + 過去形", "I wish I had more free time.", "感情を自然に伝える"),
        ("助動詞", "過去への推量・後悔", "should/might/must have p.p.", "I should have left earlier.", "後悔・推量のニュアンス"),
        ("助動詞", "丁寧な提案・意見", "should / could / might", "The government could invest more in trains.", "Part 3 で断定を避けつつ意見を述べる"),
        ("複雑な文構造", "関係詞（非制限）", ", which ~", "I missed the train, which was annoying.", "文をつなげて流暢さを出す"),
        ("複雑な文構造", "分詞構文", "Having p.p., ~ / -ing, ~", "Having finished work, I went to the gym.", "書き言葉的な洗練"),
        ("複雑な文構造", "強調構文", "What ~ is / It is ~ that", "What I enjoy most is the quiet atmosphere.", "ポイントを際立たせる"),
        ("複雑な文構造", "倒置", "Not only do ~, but ~ also", "Not only is it cheap, but it's also convenient.", "Band 8 レベルの構文"),
        ("複雑な文構造", "譲歩", "Although / Even though / While", "Although it was raining, I enjoyed the walk.", "対比でバランスのとれた議論"),
    ]
    for i, row in enumerate(rows, 2):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.font = f(j == 1)
            c.border = BOX
            c.alignment = WRAP


def main():
    wb = Workbook()
    build_guide(wb.active)
    wb.active.title = "使い方"

    example_name = "2026-09-25（例）"
    speech_ref = build_day_sheet(wb.create_sheet(example_name), EXAMPLE)
    build_day_sheet(wb.create_sheet("テンプレート"))
    build_log(wb.create_sheet("記録一覧"), example_name, speech_ref)
    build_vocab(wb.create_sheet("単語帳"), EXAMPLE["vocab"])
    build_grammar_ref(wb.create_sheet("文法リファレンス"))

    wb.save(OUT)
    print(f"saved: {OUT}")


if __name__ == "__main__":
    main()
