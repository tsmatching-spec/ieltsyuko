# IELTSリーディング間違い分析シート(v2)を生成する。使い方: python make_reading_analysis_sheet.py 出力先.xlsx
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.comments import Comment
import datetime as dt

OUT = sys.argv[1]
FONT = "Arial"
def F(**k): return Font(name=FONT, **k)
NAVY = "1F3864"; LIGHT = "D9E2F3"; INPUT = "FFF9DB"; RED = "FDE2E2"; GRAY = "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", fgColor=c)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

CATS = {
 "語彙・文法理解の問題": ["単語の意味が分からなかった","多義語の誤認識","同義語・言い換え表現に気づけなかった",
   "文法構造が複雑で理解できなかった","代名詞・指示語の指す内容を誤解","接続詞・ディスコースマーカーの見落とし"],
 "情報検索・スキャニングの問題": ["該当箇所を見つけられなかった","キーワードのパラフレーズに気づけず探せなかった",
   "似た情報と混同した","時間が足りず該当部分まで到達できなかった","複数パラグラフにまたがる情報を統合できなかった"],
 "論理的思考・推論の問題": ["因果関係の読み取りミス","比較表現の理解不足","筆者の主張と例示の区別ができなかった",
   "NOT GIVEN判定の誤り","否定表現の見落とし","限定表現の見落とし"],
 "注意力・正確性の問題": ["設問の指示を読み間違えた","答えを書き写す際のスペルミス","単数形・複数形の間違い",
   "大文字・小文字の間違い","ハイフン・スペースの有無の間違い","解答欄のずれ","時間切れで未解答"],
}
CAT_LIST = list(CATS)
DETAILS = [d for v in CATS.values() for d in v]
QTYPES = ["TRUE / FALSE / NOT GIVEN","YES / NO / NOT GIVEN","見出し選択 (Headings)","情報の場所 (Matching Information)",
  "特徴の組み合わせ (Matching Features)","文の後半選択 (Sentence Endings)","多肢選択 (Multiple Choice)",
  "文完成 (Sentence Completion)","要約完成 (Summary Completion)","表・ノート・フローチャート完成","図のラベル (Diagram)","短答式 (Short Answer)"]
BANDS = [(0,"2.5"),(4,"3.0"),(6,"3.5"),(8,"4.0"),(10,"4.5"),(13,"5.0"),(15,"5.5"),(19,"6.0"),(23,"6.5"),
         (27,"7.0"),(30,"7.5"),(33,"8.0"),(35,"8.5"),(37,"9.0")]
# Academic band thresholds as commonly published (lower bound of raw score range)
BANDS = [(0,"2.5"),(4,"3.0"),(6,"3.5"),(8,"4.0"),(10,"4.5"),(13,"5.0"),(15,"5.5"),(19,"6.0"),(23,"6.5"),
         (27,"7.0"),(30,"7.5"),(33,"8.0"),(35,"8.5"),(37,"9.0")]
BANDS = [(0,"〜3.5"),(10,"4.0"),(13,"4.5"),(15,"5.0"),(19,"5.5"),(23,"6.0"),(27,"6.5"),(30,"7.0"),(33,"7.5"),(35,"8.0"),(37,"8.5"),(39,"9.0")]

wb = Workbook()

# ---------- 選択肢マスタ ----------
m = wb.active; m.title = "選択肢マスタ"
m["A1"] = "選択肢マスタ（プルダウンの元データ。項目を変える時はここを編集）"; m["A1"].font = F(bold=True, size=12, color=NAVY)
heads = ["大カテゴリー","詳細理由（全て）","設問タイプ","正解/不正解","正解数(下限)","バンド目安(Academic)"]
for i,h in enumerate(heads):
    c = m.cell(row=3, column=i+1, value=h); c.font=F(bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=CENTER; c.border=BOX
cols = [CAT_LIST, DETAILS, QTYPES, ["正解","不正解"]]
for ci,vals in enumerate(cols):
    for ri,v in enumerate(vals):
        c = m.cell(row=4+ri, column=ci+1, value=v); c.font=F(); c.border=BOX
for ri,(n,b) in enumerate(BANDS):
    m.cell(row=4+ri, column=5, value=n).font=F(); m.cell(row=4+ri, column=6, value=b).font=F()
    m.cell(row=4+ri, column=5).border=BOX; m.cell(row=4+ri, column=6).border=BOX
m.cell(row=4+len(BANDS)+1, column=5, value="※バンド目安はAcademicの一般的な換算表に基づく目安です。回によって多少前後します。").font=F(italic=True,size=9,color="808080")
# detail -> category map for reference
m["H3"] = "詳細理由"; m["I3"] = "所属カテゴリー"
for c in (m["H3"], m["I3"]): c.font=F(bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=CENTER; c.border=BOX
r=4
for cat,ds in CATS.items():
    for d in ds:
        m.cell(row=r,column=8,value=d).font=F(); m.cell(row=r,column=9,value=cat).font=F()
        m.cell(row=r,column=8).border=BOX; m.cell(row=r,column=9).border=BOX; r+=1
for col,w in zip("ABCDEFGHI",[30,44,36,12,12,18,3,44,30]): m.column_dimensions[col].width=w

n_cat, n_det, n_q = len(CAT_LIST), len(DETAILS), len(QTYPES)
CAT_RNG = f"選択肢マスタ!$A$4:$A${3+n_cat}"
DET_RNG = f"選択肢マスタ!$B$4:$B${3+n_det}"
Q_RNG   = f"選択肢マスタ!$C$4:$C${3+n_q}"
OK_RNG  = "選択肢マスタ!$D$4:$D$5"
BAND_N  = f"選択肢マスタ!$E$4:$E${3+len(BANDS)}"
BAND_B  = f"選択肢マスタ!$F$4:$F${3+len(BANDS)}"

Q0, Q1 = 10, 49   # question rows
def passage(q): return 1 if q<=13 else (2 if q<=26 else 3)

def build_test(ws, title_no, example=None):
    for col,w in zip("ABCDEFGH",[9,9,34,12,28,38,42,42]): ws.column_dimensions[col].width=w
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = f"IELTSリーディング 間違い分析シート（{title_no}）"
    ws["A1"].font=F(bold=True,size=16,color="FFFFFF"); ws["A1"].fill=fill(NAVY); ws["A1"].alignment=CENTER
    ws.row_dimensions[1].height=32
    ws.merge_cells("A2:H2")
    ws["A2"] = "使い方：黄色のセルに入力 → 正解/不正解・大カテゴリー・詳細理由はプルダウンで選択。集計と円グラフは自動で更新されます。"
    ws["A2"].font=F(size=9,color="595959"); ws["A2"].alignment=LEFT

    # --- テスト情報 ---
    labels = [("テスト名",None),("解いた日","yyyy/mm/dd"),("分析日","yyyy/mm/dd"),("解答時間（分）",None)]
    for i,(lab,fmt) in enumerate(labels):
        r = 4+i
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
        c=ws.cell(row=r,column=1,value=lab); c.font=F(bold=True); c.fill=fill(LIGHT); c.alignment=CENTER
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        v=ws.cell(row=r,column=3); v.fill=fill(INPUT); v.font=F(color="0000FF"); v.alignment=LEFT
        if fmt: v.number_format=fmt
        for cc in range(1,5): ws.cell(row=r,column=cc).border=BOX
    ws["C4"].comment = Comment("例: Cambridge 18 Test 1", "coach")
    # score block
    ws["F4"]="正解数"; ws["G4"]=f'=COUNTIF(D{Q0}:D{Q1},"正解")&" / 40"'
    ws["F5"]="バンド目安（Academic）"; ws["G5"]=f'=IF(COUNTA(D{Q0}:D{Q1})=0,"",LOOKUP(COUNTIF(D{Q0}:D{Q1},"正解"),{BAND_N},{BAND_B}))'
    ws["F6"]="Passage別 正解数"
    ws["G6"]=(f'="P1: "&COUNTIFS(B{Q0}:B{Q1},1,D{Q0}:D{Q1},"正解")&"/13   P2: "&COUNTIFS(B{Q0}:B{Q1},2,D{Q0}:D{Q1},"正解")'
              f'&"/13   P3: "&COUNTIFS(B{Q0}:B{Q1},3,D{Q0}:D{Q1},"正解")&"/14"')
    ws["F7"]="不正解数"; ws["G7"]=f'=COUNTIF(D{Q0}:D{Q1},"不正解")'
    for r in range(4,8):
        ws.cell(row=r,column=6).font=F(bold=True); ws.cell(row=r,column=6).fill=fill(LIGHT); ws.cell(row=r,column=6).alignment=CENTER
        ws.cell(row=r,column=7).font=F(bold=True,size=11); ws.cell(row=r,column=7).alignment=LEFT
        ws.cell(row=r,column=6).border=BOX; ws.cell(row=r,column=7).border=BOX

    # --- 問題表 ---
    hdr = ["問題番号","Passage","設問タイプ","正解/不正解","大カテゴリー","詳細理由","具体的に間違いが起きた箇所","同じミスを繰り返さないために"]
    for i,h in enumerate(hdr):
        c=ws.cell(row=Q0-1,column=i+1,value=h); c.font=F(bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=CENTER; c.border=BOX
    ws.row_dimensions[Q0-1].height=28
    for q in range(1,41):
        r=Q0+q-1
        ws.cell(row=r,column=1,value=q); ws.cell(row=r,column=2,value=passage(q))
        for cc in range(1,9):
            c=ws.cell(row=r,column=cc); c.border=BOX; c.font=F(color="0000FF" if cc>2 else "000000")
            c.alignment = CENTER if cc<=4 else LEFT
            if cc>2: c.fill=fill(INPUT)
        ws.row_dimensions[r].height=30
    dvs = [(Q_RNG,"C"),(OK_RNG,"D"),(CAT_RNG,"E"),(DET_RNG,"F")]
    for rng,col in dvs:
        dv=DataValidation(type="list",formula1=f"={rng}",allow_blank=True,showErrorMessage=False)
        ws.add_data_validation(dv); dv.add(f"{col}{Q0}:{col}{Q1}")
    ws.conditional_formatting.add(f"A{Q0}:H{Q1}",FormulaRule(formula=[f'$D{Q0}="不正解"'],fill=fill(RED)))
    ws.conditional_formatting.add(f"A{Q0}:H{Q1}",FormulaRule(formula=[f'$D{Q0}="正解"'],fill=fill("E2F0D9")))
    ws.freeze_panes = f"A{Q0}"

    # --- 原因分析 ---
    s = Q1+2
    ws.merge_cells(start_row=s,start_column=1,end_row=s,end_column=8)
    c=ws.cell(row=s,column=1,value="■ 何で間違えたのか（不正解の原因の割合）"); c.font=F(bold=True,size=13,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=LEFT
    ws.row_dimensions[s].height=26
    h=s+1
    def table_header(r, first):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
        for col,txt in ((1,first),(4,"件数"),(5,"割合")):
            c=ws.cell(row=r,column=col,value=txt); c.font=F(bold=True); c.fill=fill(LIGHT); c.alignment=CENTER
        for cc in range(1,6): ws.cell(row=r,column=cc).border=BOX
    table_header(h,"大カテゴリー")
    cat_rows=[]
    items = CAT_LIST+["（カテゴリー未入力）"]
    for i,cat in enumerate(items):
        r=h+1+i; cat_rows.append(r)
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
        ws.cell(row=r,column=1,value=cat).alignment=LEFT
        crit = f'""' if i==len(CAT_LIST) else f"A{r}"
        ws.cell(row=r,column=4,value=f'=COUNTIFS($D${Q0}:$D${Q1},"不正解",$E${Q0}:$E${Q1},{crit})').alignment=CENTER
        for cc in range(1,6): ws.cell(row=r,column=cc).border=BOX; ws.cell(row=r,column=cc).font=F()
    tot=h+1+len(items)
    ws.merge_cells(start_row=tot,start_column=1,end_row=tot,end_column=3)
    ws.cell(row=tot,column=1,value="合計（不正解数）").font=F(bold=True)
    ws.cell(row=tot,column=4,value=f"=SUM(D{cat_rows[0]}:D{cat_rows[-1]})").font=F(bold=True)
    ws.cell(row=tot,column=4).alignment=CENTER
    for cc in range(1,6): ws.cell(row=tot,column=cc).border=BOX; ws.cell(row=tot,column=cc).fill=fill(GRAY)
    for r in cat_rows+[tot]:
        c=ws.cell(row=r,column=5,value=f'=IF($D${tot}=0,"",D{r}/$D${tot})'); c.number_format="0%"; c.alignment=CENTER
    ws.conditional_formatting.add(f"D{cat_rows[0]}:D{cat_rows[-1]}",DataBarRule(start_type="num",start_value=0,end_type="max",color="4472C4"))

    pie=PieChart(); pie.title="大カテゴリー別 間違いの割合"
    pie.add_data(Reference(ws,min_col=4,min_row=h,max_row=cat_rows[-1]),titles_from_data=True)
    pie.set_categories(Reference(ws,min_col=1,min_row=cat_rows[0],max_row=cat_rows[-1]))
    pie.dataLabels=DataLabelList(); pie.dataLabels.showPercent=True; pie.dataLabels.showVal=False
    pie.dataLabels.showCatName=False; pie.dataLabels.showSerName=False; pie.dataLabels.showLeaderLines=True
    colors=["4472C4","ED7D31","70AD47","FFC000","A5A5A5"]
    ser=pie.series[0]
    for i,col in enumerate(colors):
        pt=DataPoint(idx=i); pt.graphicalProperties.solidFill=col; ser.dPt.append(pt)
    pie.legend.position="r"; pie.height=8.5; pie.width=15
    ws.add_chart(pie, f"F{h}")

    # detail table
    d0 = max(tot+2, h+18)
    table_header(d0,"詳細理由")
    ws.cell(row=d0,column=6,value="カテゴリー").font=F(bold=True)
    ws.cell(row=d0,column=6).fill=fill(LIGHT); ws.cell(row=d0,column=6).alignment=CENTER; ws.cell(row=d0,column=6).border=BOX
    det_rows=[]
    for i,d in enumerate(DETAILS):
        r=d0+1+i; det_rows.append(r)
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
        ws.cell(row=r,column=1,value=d).alignment=LEFT
        cat=[k for k,v in CATS.items() if d in v][0]
        ws.cell(row=r,column=6,value=cat).font=F(size=9,color="595959")
        ws.cell(row=r,column=4,value=f'=COUNTIFS($D${Q0}:$D${Q1},"不正解",$F${Q0}:$F${Q1},A{r})').alignment=CENTER
        c=ws.cell(row=r,column=5,value=f'=IF($D${tot}=0,"",D{r}/$D${tot})'); c.number_format="0%"; c.alignment=CENTER
        for cc in range(1,7): ws.cell(row=r,column=cc).border=BOX
        for cc in range(1,6): ws.cell(row=r,column=cc).font=F()
    ws.conditional_formatting.add(f"D{det_rows[0]}:D{det_rows[-1]}",DataBarRule(start_type="num",start_value=0,end_type="max",color="ED7D31"))
    ws.conditional_formatting.add(f"A{det_rows[0]}:E{det_rows[-1]}",
        FormulaRule(formula=[f'AND($D{det_rows[0]}>0,$D{det_rows[0]}=MAX($D${det_rows[0]}:$D${det_rows[-1]}))'],font=Font(bold=True,color="C00000")))
    note=det_rows[-1]+1
    ws.merge_cells(start_row=note,start_column=1,end_row=note,end_column=8)
    ws.cell(row=note,column=1,value="※ 最も多い詳細理由は赤字で表示されます。").font=F(size=9,color="808080")

    # --- 振り返り ---
    p = note+2
    ws.merge_cells(start_row=p,start_column=1,end_row=p,end_column=8)
    c=ws.cell(row=p,column=1,value="■ 振り返りと今後の対策"); c.font=F(bold=True,size=13,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=LEFT
    ws.row_dimensions[p].height=26
    blocks=[("間違えた傾向（自分で書いてみよう）","上の円グラフ・表を見て、どんな間違いが多かったか、どの設問タイプ・Passageで間違えやすいかを書きましょう。",5),
            ("解決するために、これから対策すること","例：パラフレーズ表現を毎日10個ノートにまとめる／T/F/NGを1日1セット解く など",5),
            ("今後、問題を解くときに意識すること","例：NOT GIVENは「本文に書いていない」を根拠にする／転記後に単数・複数を確認する など",4)]
    r=p+1
    refl={}
    for lab,hint,n in blocks:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
        c=ws.cell(row=r,column=1,value=lab); c.font=F(bold=True); c.fill=fill(LIGHT); c.alignment=LEFT
        for cc in range(1,9): ws.cell(row=r,column=cc).border=BOX
        ws.merge_cells(start_row=r+1,start_column=1,end_row=r+n,end_column=8)
        c=ws.cell(row=r+1,column=1); c.fill=fill(INPUT); c.alignment=TOP; c.font=F(color="0000FF")
        c.comment=Comment(hint,"coach")
        for rr in range(r+1,r+n+1):
            for cc in range(1,9): ws.cell(row=rr,column=cc).border=BOX
        refl[lab]=r+1
        r+=n+2
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    c=ws.cell(row=r,column=1,value="次回の問題を解く日"); c.font=F(bold=True); c.fill=fill(LIGHT); c.alignment=CENTER
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
    v=ws.cell(row=r,column=3); v.fill=fill(INPUT); v.number_format="yyyy/mm/dd"; v.font=F(color="0000FF",bold=True); v.alignment=CENTER
    ws.merge_cells(start_row=r,start_column=5,end_row=r,end_column=6)
    ws.cell(row=r,column=5,value="次回解く予定の問題").font=F(bold=True); ws.cell(row=r,column=5).fill=fill(LIGHT); ws.cell(row=r,column=5).alignment=CENTER
    ws.merge_cells(start_row=r,start_column=7,end_row=r,end_column=8)
    ws.cell(row=r,column=7).fill=fill(INPUT); ws.cell(row=r,column=7).font=F(color="0000FF"); ws.cell(row=r,column=7).alignment=LEFT
    for cc in range(1,9): ws.cell(row=r,column=cc).border=BOX
    ws.row_dimensions[r].height=24
    ws.print_area=f"A1:H{r}"
    ws.page_setup.orientation="portrait"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.sheet_properties.tabColor = "FFC000" if example else "4472C4"
    info = dict(next_row=r, refl=refl, cat_rows=cat_rows)

    if example:
        ws["C4"]="Cambridge 18 Test 1"; ws["C5"]=dt.date(2026,10,20); ws["C6"]=dt.date(2026,10,21); ws["C7"]=62
        wrong = example["wrong"]
        for q in range(1,41):
            rr=Q0+q-1
            ws.cell(row=rr,column=3,value=example["types"][q])
            if q in wrong:
                cat,det,where,fix = wrong[q]
                ws.cell(row=rr,column=4,value="不正解"); ws.cell(row=rr,column=5,value=cat); ws.cell(row=rr,column=6,value=det)
                ws.cell(row=rr,column=7,value=where); ws.cell(row=rr,column=8,value=fix)
            else:
                ws.cell(row=rr,column=4,value="正解")
        R=refl
        ws.cell(row=R[blocks[0][0]],column=1,value=example["trend"])
        ws.cell(row=R[blocks[1][0]],column=1,value=example["plan"])
        ws.cell(row=R[blocks[2][0]],column=1,value=example["mind"])
        ws.cell(row=r,column=3,value=dt.date(2026,10,27)); ws.cell(row=r,column=7,value="Cambridge 18 Test 2")
    return info

L,S,R_,A = CAT_LIST
types={}
for q in range(1,41):
    types[q] = (QTYPES[0] if q<=7 else QTYPES[9] if q<=13 else QTYPES[2] if q<=20 else QTYPES[1] if q<=26
                else QTYPES[6] if q<=31 else QTYPES[4] if q<=35 else QTYPES[8])
example = dict(types=types, wrong={
 3:(R_,"NOT GIVEN判定の誤り","本文に近い記述があったのでTRUEにしたが、頻度の情報は書かれていなかった","本文にない情報を自分の知識で補っていないか確認する"),
 6:(R_,"限定表現の見落とし","問題文の only / always を見落とした","only・all・always などの限定語に○をつける"),
 10:(A,"単数形・複数形の間違い","正解は seeds なのに seed と書いた","転記後に冠詞・複数形を本文と照合する"),
 15:(L,"同義語・言い換え表現に気づけなかった","decline が decrease の言い換えだと気づかなかった","パラフレーズノートに追加して復習"),
 17:(S,"該当箇所を見つけられなかった","段落Dの主旨をつかめず見出しを決められなかった","各段落の1文目と最後の文を先に読む"),
 19:(L,"同義語・言い換え表現に気づけなかった","feasible = possible と結びつかなかった","パラフレーズノートに追加して復習"),
 22:(R_,"NOT GIVEN判定の誤り","筆者の意見が書かれていないのにNOにした","NO＝本文と逆の主張がある時だけ、と確認する"),
 25:(R_,"否定表現の見落とし","hardly を見落として肯定文として読んだ","否定語（hardly, rarely, few）に線を引く"),
 28:(L,"文法構造が複雑で理解できなかった","長い関係詞節で主語と動詞を取り違えた","SVを先に見つけてから修飾部分を読む"),
 33:(S,"似た情報と混同した","研究者Bの意見を研究者Cのものと混同した","人名に□をつけて意見を整理しながら読む"),
 38:(A,"時間切れで未解答","Passage3に入った時点で残り12分しかなかった","P1を17分以内で終える時間配分を練習する"),
 39:(A,"時間切れで未解答","同上","同上"),
},
 trend="論理的思考・推論の問題（特にNOT GIVENの判定）が一番多い。T/F/NG と Y/N/NG で、本文に近い記述があると TRUE/YES にしてしまう。Passage3 は時間切れで2問落とした。",
 plan="① T/F/NG・Y/N/NG を毎日1セット解き、「根拠の一文」を本文から書き出す\n② 間違えた言い換えをパラフレーズノートにまとめ、週末に見直す\n③ P1を17分・P2を20分で解く時間配分トレーニングをする",
 mind="・NOT GIVEN は「本文に書かれていない」ことを根拠に選ぶ\n・only / always / hardly など限定・否定の語に印をつける\n・解答を書き写したら単数・複数をチェックする")

ex = wb.create_sheet("記入例")
build_test(ex,"記入例",example)
N_TESTS=5
infos={}
for i in range(1,N_TESTS+1):
    ws=wb.create_sheet(f"テスト{i}")
    infos[ws.title]=build_test(ws,f"テスト{i}")

# ---------- テスト一覧 ----------
sm = wb.create_sheet("テスト一覧",0)
sm.sheet_view.showGridLines=False
sm.merge_cells("A1:K1"); sm["A1"]="IELTSリーディング テスト一覧（各テストシートから自動で集計）"
sm["A1"].font=F(bold=True,size=16,color="FFFFFF"); sm["A1"].fill=fill(NAVY); sm["A1"].alignment=CENTER; sm.row_dimensions[1].height=32
sm.merge_cells("A2:K2"); sm["A2"]="このシートは入力不要です。各「テスト」シートに記入すると自動で反映されます。テストを増やす時は「テスト5」シートをコピーし、この表に1行追加してください。"
sm["A2"].font=F(size=9,color="595959"); sm["A2"].alignment=LEFT
hd=["シート","テスト名","解いた日","分析日","正解数","バンド目安","語彙・文法","情報検索","論理・推論","注意力・正確性","次回の問題を解く日"]
for i,t in enumerate(hd):
    c=sm.cell(row=4,column=i+1,value=t); c.font=F(bold=True,color="FFFFFF"); c.fill=fill(NAVY); c.alignment=CENTER; c.border=BOX
sm.row_dimensions[4].height=30
names=["記入例"]+[f"テスト{i}" for i in range(1,N_TESTS+1)]
ref_info = infos["テスト1"]
for j,nm in enumerate(names):
    r=5+j; q=f"'{nm}'!"
    inf = ref_info  # all test sheets share identical layout
    vals=[nm, f'=IF({q}C4="","",{q}C4)', f'=IF({q}C5="","",{q}C5)', f'=IF({q}C6="","",{q}C6)',
          f'=IF(COUNTA({q}D{Q0}:D{Q1})=0,"",COUNTIF({q}D{Q0}:D{Q1},"正解"))', f'={q}G5']
    vals += [f"={q}D{inf['cat_rows'][k]}" for k in range(4)]
    vals += [f'=IF({q}C{inf["next_row"]}="","",{q}C{inf["next_row"]})']
    for i,v in enumerate(vals):
        c=sm.cell(row=r,column=i+1,value=v); c.font=F(); c.border=BOX; c.alignment=CENTER
        if i in (2,3,10): c.number_format="yyyy/mm/dd"
    sm.cell(row=r,column=1).hyperlink=f"#'{nm}'!A1"; sm.cell(row=r,column=1).font=F(color="0563C1",underline="single")
    if nm=="記入例":
        for i in range(11): sm.cell(row=r,column=i+1).fill=fill(GRAY)
for col,w in zip("ABCDEFGHIJK",[12,28,13,13,9,11,11,11,11,13,16]): sm.column_dimensions[col].width=w
sm.cell(row=6+N_TESTS,column=1,value="※ 語彙・文法〜注意力・正確性の列は、各テストの不正解の件数です。").font=F(size=9,color="808080")
sm.sheet_properties.tabColor="1F3864"
sm.page_setup.orientation="landscape"; sm.sheet_properties.pageSetUpPr.fitToPage=True; sm.page_setup.fitToWidth=1; sm.page_setup.fitToHeight=0
sm.freeze_panes="A5"

# ---------- 使い方 ----------
hw = wb.create_sheet("使い方", 0)
hw.sheet_view.showGridLines = False
hw.column_dimensions["A"].width = 3
hw.column_dimensions["B"].width = 100
hw.merge_cells("A1:B1"); hw["A1"] = "IELTSリーディング 間違い分析シート ― 使い方"
hw["A1"].font=F(bold=True,size=16,color="FFFFFF"); hw["A1"].fill=fill(NAVY); hw["A1"].alignment=CENTER
hw.row_dimensions[1].height=36
sections = [
 ("■ はじめに：必ずコピーしてから使ってください", [
   "このファイルは配布用の原本です。直接書き込まず、必ずご自身用のコピーを作成してから使ってください。",
   "・Googleスプレッドシートの場合：メニューの「ファイル」→「コピーを作成」→ 名前を付けて自分のドライブに保存",
   "・Excelの場合：「ファイル」→「名前を付けて保存」で別名のファイルとして保存",
   "・ファイル名の例：「（お名前）_リーディング間違い分析シート」",
 ], "FFF2CC"),
 ("■ シートの構成", [
   "・使い方 ………… このページです",
   "・テスト一覧 …… 各テストの結果が自動でまとまります（入力不要）",
   "・記入例 ………… 記入の見本です。まずはこのシートで完成イメージを確認してください",
   "・テスト1〜5 …… 1回のテストにつき1シートを使って記入します",
   "・選択肢マスタ … プルダウンの元データです（編集しないでください）",
 ], None),
 ("■ 記入の手順", [
   "① 問題を解いたら、空いている「テスト」シートを開き、テスト名・解いた日・分析日・解答時間を入力する",
   "② 答え合わせをして、40問すべての「正解/不正解」をプルダウンで選ぶ（設問タイプも選ぶ）",
   "③ 不正解の問題について「大カテゴリー」「詳細理由」をプルダウンで選び、「具体的に間違いが起きた箇所」「同じミスを繰り返さないために」を書く",
   "④ 表の下の円グラフ・集計表で、何で間違えたのかの割合を確認する（自動で更新されます）",
   "⑤ 一番下の「振り返りと今後の対策」に、間違えた傾向・これから対策すること・今後意識すること・次回の問題を解く日を書く",
   "※ 入力するのは黄色のセルだけです。白いセル（自動計算部分）は消したり書き換えたりしないでください。",
   "※ テストが5回分を超えたら「テスト5」シートをコピーして使ってください。",
 ], None),
 ("■ 外部転用禁止", [
   "本シートは受講生ご本人の学習用としてのみご利用いただけます。",
   "・本シートおよびその内容（構成・項目・分析カテゴリー等）の無断での転載、複製、再配布、販売を禁止します。",
   "・第三者への共有・譲渡、SNSやブログ等への掲載、他の講座・教材への転用はご遠慮ください。",
   "・上記に違反する行為が確認された場合、利用を停止させていただくことがあります。",
 ], "FDE2E2"),
]
r = 3
for head, lines, bg in sections:
    c = hw.cell(row=r, column=2, value=head); c.font=F(bold=True,size=13,color=NAVY)
    c.border = Border(bottom=Side(style="medium", color=NAVY)); r += 1
    for ln in lines:
        c = hw.cell(row=r, column=2, value=ln); c.font=F(size=11); c.alignment=LEFT
        if bg: c.fill = fill(bg)
        hw.row_dimensions[r].height = 22 if len(ln) < 48 else 36; r += 1
    r += 1
hw.sheet_properties.tabColor = "C00000"
hw.page_setup.orientation="portrait"; hw.sheet_properties.pageSetUpPr.fitToPage=True
hw.page_setup.fitToWidth=1; hw.page_setup.fitToHeight=0

# move master to end
wb.move_sheet("選択肢マスタ", offset=len(wb.sheetnames))
wb.active = 0
wb.save(OUT)
print("saved", OUT)
