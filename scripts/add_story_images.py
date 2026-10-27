"""切り出したストーリーズ画像を月別シートの各日の欄に横に並べて貼り、内容メモを記入する。"""
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

BOOK = "sns_tracking/SNS記録シート.xlsx"
SHEET = "10月ストーリーズ"
IMG_DIR = "sns_tracking/stories/2026-10"
FIRST_ROW = 5
ROWS_PER_DAY = 3
SLOT_COL = 6  # F列から横に並べる
IMG_HEIGHT = 255  # px（行の高さ210pt=280pxに余白付きで収まるサイズ）
PAD = 8  # px（セルの左上からの余白）

# (日, その日の何枚目か, 内容メモ) を投稿順に。画像は IMG_DIR/10-DD_NN.png
STORIES = [
    (1, 1, "ポッドキャスト更新 #3 負け癖からの脱却"),
    (1, 2, "YouTubeに届いたコメント（社会人は時間がないから伸びない？→答えはNO）"),
    (1, 3, "個別学習相談に来られる方々からの声（つまり…）"),
    (1, 4, "正しい方向に最短で進める仕組みが必要／IELTS個別学習相談リンク"),
    (2, 1, "今週早速一歩踏み出したみなさま（相談申込みフォーム通知）"),
    (2, 2, "完璧を待っていたら何も始まらない／「気になる」DM誘導"),
    (3, 1, "惜しい!! ライティングの勉強法を変えたら伸びる人"),
    (3, 2, "評価基準に沿った書き方を知らないから"),
    (3, 3, "頑張りどころを間違えなければ／書き方の方法知りたい？（😍スタンプ）"),
    (4, 1, "昨日のリアクションありがとう→3ステップで解説"),
    (4, 2, "ステップ1：評価基準に沿ったライティング教材を選ぶ"),
    (4, 3, "ステップ2：タスクごとの書き方ルールをインプット"),
    (4, 4, "ステップ3：ルール確認のために書く／受講生の声"),
    (4, 5, "動画講座：全9コース・100本以上"),
    (4, 6, "1本5分以内／W5.5→6.5受講生インタビューリンク"),
    (4, 7, "ライブ告知（22時〜 タスク2「書くことがない」を解決）"),
    (5, 1, "YouTube台本は自分で作ってる（今週末公開の動画から→）"),
    (5, 2, "スライド追加で見やすく／土曜日の公開をお楽しみに"),
    (6, 1, "春にOA5.0→6.5以上を達成したい／個別学習相談が増えました"),
    (6, 2, "大学院奨学金応募に向けてラストスパート"),
    (8, 1, "「テンプレートもらえますか？」→お渡ししていません"),
    (8, 2, "テンプレート通りに書いてもスコアは上がらないから"),
    (8, 3, "受講生の例：テンプレをやめて3ヶ月で6.0→6.5"),
    (8, 4, "半年後、差がつく人とつかない人"),
    (9, 1, "テーマ関連語彙を覚えるのは一種のギャンブル（音声あり）"),
    (9, 2, "YouTube動画公開：スコアが一気に伸びる人の特徴"),
    (10, 1, "嬉しいお知らせが届きました"),
    (10, 2, "2ヶ月で3技能ベストスコア更新（高校生）"),
    (10, 3, "やる気にもつながってます／次の目標"),
    (11, 1, "こんなこと言われたら泣いてしまう"),
    (11, 2, "娘の良き先生になってくれそう（相談の感想）"),
    (11, 3, "ライブ告知（22時〜 手遅れになる前に知っておくべきこと）"),
    (12, 1, "「苦手」を観察すると本当の課題が見えてくる（相談の感想）"),
    (13, 1, "コーチング受講生はスコアアップ率95%"),
    (13, 2, "学習の姿勢が整っている人だけ（アンケート：あなたはどっち？）"),
    (13, 3, "スコアアップされた受講生さんは…"),
    (13, 4, "この両方＝学習の姿勢が整っていた人たち"),
    (13, 5, "学習迷子から卒業して行動を続けられるから"),
    (13, 6, "姿勢が定まっているなら正しく進むだけ／IELTS個別学習相談リンク"),
    (13, 7, "今日のリール：ライティングタスク1の超もったいないミス"),
    (14, 1, "孤独から脱出できました"),
    (15, 1, "2026年も残り2か月／ゴール設定が春から夏に"),
    (15, 2, "来年の春までにOA1.0以上上げないといけない人いますか？（🙌スタンプ）"),
    (16, 1, "こんなに春までにOA1.0以上アップしたい人が／何か企画をしようかな"),
    (16, 2, "勉強会・セミナーのテーマ募集（アンケート）"),
]

wb = load_workbook(BOOK)
ws = wb[SHEET]
for day, seq, memo in STORIES:
    row = FIRST_ROW + (day - 1) * ROWS_PER_DAY  # その日のスクショ行（下が内容、閲覧数）
    col = SLOT_COL + seq - 1
    ws.cell(row=row + 1, column=col, value=memo)
    img = Image(f"{IMG_DIR}/10-{day:02d}_{seq:02d}.png")
    img.width, img.height = round(img.width * IMG_HEIGHT / img.height), IMG_HEIGHT
    marker = AnchorMarker(col=col - 1, colOff=pixels_to_EMU(PAD), row=row - 1, rowOff=pixels_to_EMU(PAD))
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
    ws.add_image(img)
wb.save(BOOK)
print(f"added {len(STORIES)} stories to {SHEET}")
