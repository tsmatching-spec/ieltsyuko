"""切り出したストーリーズ画像を月別シートに順番に貼り付け、投稿日・内容メモを記入する。"""
from datetime import date

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

BOOK = "sns_tracking/SNS記録シート.xlsx"
SHEET = "10月ストーリーズ"
IMG_DIR = "sns_tracking/stories/2026-10"
FIRST_ROW = 5
IMG_HEIGHT = 260  # px（行の高さ220pt≒293pxに余白付きで収まるサイズ）
PAD = 8  # px（セルの左上からの余白）

# (日, 内容メモ) を投稿順に
STORIES = [
    (1, "ポッドキャスト更新 #3 負け癖からの脱却"),
    (2, "YouTubeに届いたコメント（社会人は時間がないから伸びない？→答えはNO）"),
    (2, "個別学習相談に来られる方々からの声（つまり…）"),
    (2, "正しい方向に最短で進める仕組みが必要／IELTS個別学習相談リンク"),
    (3, "今週早速一歩踏み出したみなさま（相談申込みフォーム通知）"),
    (3, "完璧を待っていたら何も始まらない／「気になる」DM誘導"),
    (4, "惜しい!! ライティングの勉強法を変えたら伸びる人"),
    (4, "評価基準に沿った書き方を知らないから"),
    (4, "頑張りどころを間違えなければ／書き方の方法知りたい？（リアクションスタンプ）"),
    (5, "昨日のリアクションありがとう→3ステップで解説"),
    (5, "ステップ1：評価基準に沿ったライティング教材を選ぶ"),
    (5, "ステップ2：タスクごとの書き方ルールをインプット"),
    (5, "ステップ3：ルール確認のために書く／受講生の声"),
    (5, "動画講座：全9コース・100本以上"),
    (5, "1本5分以内／W5.5→6.5受講生インタビューリンク"),
    (5, "ライブ告知（日本時間22時〜 タスク2「書くことがない」を解決）"),
]

wb = load_workbook(BOOK)
ws = wb[SHEET]
ws.column_dimensions["C"].width = 24
for n, (day, memo) in enumerate(STORIES):
    row = FIRST_ROW + n
    ws.cell(row=row, column=2, value=date(2026, 10, day))
    ws.cell(row=row, column=4, value=memo)
    img = Image(f"{IMG_DIR}/{n + 1:02d}.png")
    img.width, img.height = round(img.width * IMG_HEIGHT / img.height), IMG_HEIGHT
    marker = AnchorMarker(col=2, colOff=pixels_to_EMU(PAD), row=row - 1, rowOff=pixels_to_EMU(PAD))
    img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
    ws.add_image(img)
wb.save(BOOK)
print(f"added {len(STORIES)} stories to {SHEET}")
