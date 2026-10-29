# -*- coding: utf-8 -*-
"""Nomura CIO Monthly 2026年10月号 要約デッキ（v2：原資料の写真・図表を活用）

python3 build.py <資材dir> <テンプレート.pptx> <出力.pptx> <画像dir>
画像dir には photo_*.jpg（PDF埋め込み写真）と fig*.png（図表の切り出し）を置く。
"""
import os
import sys
sys.path.insert(0, sys.argv[1])
from nomura_deck import Deck, C

TEMPLATE, OUT, IMG = sys.argv[2], sys.argv[3], sys.argv[4]
I = lambda n: os.path.join(IMG, n)
d = Deck(TEMPLATE)
R, N, G = C.RED, C.NAVY, C.GRAY

# 共通レイアウト：左に原資料の図表2枚、右に要点パネル、下に帯
def two_charts(s, a, b, cap_a, cap_b, y=0.95, h=3.35, w=2.8):
    d.image_fit(s, I(a), 0.4, y, w, h, caption=cap_a)
    d.image_fit(s, I(b), 0.4 + w + 0.15, y, w, h, caption=cap_b)

def panel(s, head, paras, x=6.35, y=0.95, w=3.25, h=3.35, color=N, size=10.5):
    d.box(s, x, y, w, h, fill=C.BG)
    d.text(s, x + 0.15, y + 0.08, w - 0.3, 0.35, head, size=12, bold=True, color=color)
    d.text(s, x + 0.15, y + 0.48, w - 0.3, h - 0.55, paras, size=size, bullet=True, space_after=5)

# 1 表紙 ---------------------------------------------------------------
d.cover("Nomura CIO Monthly\n2026年10月号　要約", [
    "金利上昇・AI半導体の変調・信用リスクに備える",
    "出典：野村證券「Nomura CIO Monthly」2026年10月号",
    "投資判断：2026年9月18日時点／比較：同9月号",
    "要約作成日：2026年10月29日",
], photo=I("photo_cover.jpg"), photo_w=4.2)

# 2 結論 ---------------------------------------------------------------
s = d.slide("今月の結論：金利上昇とAI変調に備える")
cards = [
    ("金利上昇はREITの強い逆風に", ["長期金利の上昇で利回り面の優位性が低下", "先進国REITを「ややオーバー」→「ニュートラル」に引き下げ"]),
    ("AI・半導体景気の変調に備える", ["韓国の出荷・在庫バランスがマイナスに転化", "メモリ価格は頭打ち、「AI脅威論」も台頭"]),
    ("低格付け先融資で信用リスク増大", ["CCC格向けローンの上乗せ金利が上昇", "ハイイールド社債はニュートラル維持、波及なら引き下げ検討"]),
]
cw = (9.2 - 0.3 * 2) / 3
for i, (h, b) in enumerate(cards):
    d.card(s, 0.4 + i * (cw + 0.3), 0.95, cw, 2.1, h, b, num=i + 1, head_size=13, body_size=11, head_h=0.62)
d.text(s, 0.4, 3.3, 4.2, 0.3, "資産クラス別のスタンス（10月号）", size=11, bold=True, color=C.SUB)
stance = [("株式", "ややオーバー", C.SCALE[3]), ("クレジット", "ややアンダー", C.SCALE[1]),
          ("金利資産", "ニュートラル", C.SCALE[2]), ("REIT", "ニュートラル", C.SCALE[2])]
for i, (a, v, col) in enumerate(stance):
    x = 0.4 + i * 1.08
    d.box(s, x, 3.65, 1.0, 1.1, fill=C.BG)
    d.text(s, x, 3.7, 1.0, 0.3, a, size=11, bold=True, align="c")
    d.box(s, x + 0.36, 4.02, 0.28, 0.28, fill=col, shape="oval")
    d.text(s, x, 4.36, 1.0, 0.3, v, size=10, align="c", color=C.SUB)
d.text(s, 4.95, 3.3, 4.6, 0.3, "投資判断の変更履歴（直近2か月）", size=11, bold=True, color=C.SUB)
chg = [("10月号", "先進国REIT", "ややオーバー→ニュートラル", N),
       ("9月号", "新興国株式", "ニュートラル→ややアンダー", N),
       ("9月号", "先進国国債（ヘッジ有）", "ややアンダー→ニュートラル", R)]
for i, (m, a, v, col) in enumerate(chg):
    y = 3.65 + i * 0.37
    d.tag(s, 4.95, y + 0.03, 0.7, m, color=C.SUB if m == "9月号" else R, h=0.26, size=9)
    d.text(s, 5.72, y, 1.8, 0.32, a, size=10.5, bold=True, anchor="m")
    d.text(s, 7.45, y, 2.2, 0.32, [[("▼ " if col == N else "▲ ", {"color": col, "bold": True}), (v, {})]],
           size=9.5, anchor="m")
d.text(s, 4.95, 4.78, 4.6, 0.3, "※上記以外の資産は、9月号・10月号とも判断を維持", size=9.5, color=C.SUB)
d.notes(s, "出典：10月号 表紙「今月の投資戦略」、p.10 CIO VIEW、9月号 p.12 CIO VIEW")

# 3 CIO VIEW -----------------------------------------------------------
s = d.slide("CIO VIEW：先進国REITを引き下げ、他は維持")
rows = [
    dict(kind="group", name="株式", now=4, note=""),
    dict(kind="asset", name="国内株式", now=3, note="G>R・自社株買いが支え、原油高で上昇ペースは緩やか"),
    dict(kind="asset", name="先進国株式", now=4, note="米国は原油輸出国の恩恵、欧州は割安感と財政"),
    dict(kind="asset", name="新興国株式", now=2, note="AI期待の後退、中国景気の低迷、異常気象"),
    dict(kind="group", name="クレジット", now=2, note=""),
    dict(kind="asset", name="投資適格社債", now=2, note="割高感、長いデュレーションに米高金利が逆風"),
    dict(kind="asset", name="ハイイールド社債", now=3, note="割高だが高利回り需要で拡大余地は限定"),
    dict(kind="asset", name="新興国債券", now=2, note="割高感、交易条件悪化でソブリンリスク"),
    dict(kind="group", name="金利資産", now=3, note=""),
    dict(kind="asset", name="国内債券", now=2, note="積極財政を背景に金利上昇圧力が継続"),
    dict(kind="asset", name="先進国国債（ヘッジ有）", now=3, note="中銀の利上げで長期金利の上値は限定的"),
    dict(kind="group", name="REIT", now=3, note=""),
    dict(kind="asset", name="国内REIT", now=2, note="分配金利回りスプレッドが低水準で妙味薄"),
    dict(kind="asset", name="先進国REIT", now=3, prev=4, note="金利上昇で優位性低下、AI脅威論が重し"),
]
d.view_matrix(s, 0.4, 0.86, 9.2, rows, comment_w=4.0, row_h=0.26, name_w=1.85)
d.text(s, 0.4, 5.32, 6.5, 0.25, [[("◯ 点線＝前月　", {"color": C.SUB}), ("▼ 紺の矢印＝引き下げ", {"color": N}),
                                  ("　▲ 赤の矢印＝引き上げ", {"color": R})]], size=9)
d.notes(s, "今後6か月～1年程度の相対的な投資判断（2026年9月18日時点）。出典：p.10")

# 4 F/Vマップ -----------------------------------------------------------
s = d.slide("各資産の位置：今月は目立った局面変化なし")
d.image_fit(s, I("fig7.png"), 0.4, 0.95, 5.75, 4.0, caption="原資料 p.9 図表7（2026年9月14日時点、点線＝前月）")
x0, y0 = 6.4, 0.95
d.text(s, x0, y0, 3.2, 0.3, "マップの読み方", size=12, bold=True, color=N)
qs = [("割安×悪化", C.BG, C.SUB), ("割安×改善", C.RED_T, R), ("割高×悪化", C.NAVY_T, N), ("割高×改善", C.BG, C.SUB)]
for i, (lab, f, col) in enumerate(qs):
    d.box(s, x0 + (i % 2) * 1.6, y0 + 0.38 + (i // 2) * 0.62, 1.55, 0.57, fill=f, text=lab, size=10.5, bold=True,
          color=col, shape="rect")
d.text(s, x0, y0 + 1.7, 3.2, 0.5, "縦軸＝バリュエーション（上ほど割安）、横軸＝ファンダメンタルズ（右ほど改善）", size=9.5,
       color=C.SUB)
d.text(s, x0, y0 + 2.3, 3.2, 0.3, "今月のポイント", size=12, bold=True, color=N)
d.text(s, x0, y0 + 2.65, 3.2, 1.4, [
    "どの資産も局面の変化はなし",
    "2025年以降、バリュエーション指標の有効性は低下",
    "円高リスクに備え、割高圏の先進国REITを引き下げ"], size=10.5, bullet=True, space_after=4)
d.callout(s, 0.4, 4.75 + 0.2, 9.2, 0.45, "CIOの姿勢", "不透明な環境でも、長期的にリスク資産を保有することを重視", color=N)
d.notes(s, "出典：p.9 ファンダメンタルズとバリュエーションの目立った変化なし")

# 5 地域別 -------------------------------------------------------------
s = d.slide("経済見通し：米国は堅調、日欧は物価高に警戒")
cols = [
    ("日本", "photo_japan.jpg", [("日本", "まだら", G, ["設備投資が牽引、消費は鈍化懸念", "日銀は9月に利上げ（前回から3か月）",
                                                    "CIO予想：27年1月・4月も利上げ", "原油高による交易条件の悪化に注意"])]),
    ("先進国", "photo_developed.jpg", [("米国", "堅調", R, ["8月雇用者数 +16.2万人、コアCPI 前月比+0.4%",
                                                         "FRBは9月に約3年ぶり利上げ、12月も予想"]),
                                     ("ユーロ圏", "警戒", G, ["4-6月GDP 前期比年率+1.8%",
                                                          "エネルギー高・渇水が重し、12月も利上げ想定"])]),
    ("新興国", "photo_emerging.jpg", [("中国", "低迷", N, ["8月小売売上高 前年比+0.4%", "政府の景気支援に本気度乏しい"]),
                                    ("中国以外", "警戒", G, ["半導体市況の鈍化、原油高・エルニーニョ", "インドで利上げ観測"])]),
]
cw = (9.2 - 0.25 * 2) / 3
for i, (nm, ph, subs) in enumerate(cols):
    x = 0.4 + i * (cw + 0.25)
    d.box(s, x, 0.95, cw, 4.4, fill=C.BG, shape="rect")
    d.image_fill(s, I(ph), x, 0.95, cw, 1.65)
    d.box(s, x, 2.24, 1.15, 0.36, fill=C.TEXT, shape="rect", text=nm, size=13, bold=True, color=C.WHITE)
    y = 2.72
    for (sn, tone, col, pts) in subs:
        d.text(s, x + 0.12, y, 1.3, 0.3, sn, size=12, bold=True, anchor="m")
        d.tag(s, x + cw - 0.82, y + 0.02, 0.7, tone, color=col, h=0.26, size=9.5)
        hh = 0.45 * len(pts) + 0.1
        d.text(s, x + 0.12, y + 0.36, cw - 0.24, hh, pts, size=10.5, bullet=True, space_after=3)
        y += 0.36 + hh + 0.12
d.notes(s, "写真・本文の出典：p.2 各地域の足元の経済環境と今後の見通し。トーン分類は要約者による。")

# 6 金利 ---------------------------------------------------------------
s = d.slide("金利：利上げ再開で世界的に長期金利が上昇")
two_charts(s, "fig2_1.png", "fig2_2.png", "原資料 p.4 図表2-1", "原資料 p.4 図表2-2", h=3.3)
panel(s, "ポイント", [
    "10年金利は米国5%台、ドイツ3%台半ば、日本3%近辺",
    "主因はインフレ懸念：3月末の米・イラン衝突で原油が急騰し、各中銀が利上げ",
    "日本は上昇ペースが最速。日銀の国債保有の縮小がFRB・ECBより速い（右図）"], h=3.3)
evs = [("2026年9月", "日銀・FRB・ECBが利上げ", R), ("2026年12月", "FRB・ECBが利上げ（予想）", G),
       ("2027年1月・4月", "日銀が利上げ（予想）", G)]
d.text(s, 0.4, 4.4, 2, 0.3, "利上げの実績とCIO予想", size=10.5, bold=True, color=C.SUB)
ew = (9.2 - 0.2 * 2) / 3
for i, (w_, t, col) in enumerate(evs):
    x = 0.4 + i * (ew + 0.2)
    d.box(s, x, 4.7, ew, 0.62, fill=C.RED_T if col == R else C.BG)
    d.text(s, x + 0.1, 4.72, ew - 0.2, 0.28, w_, size=10.5, bold=True, color=col if col == R else C.TEXT)
    d.text(s, x + 0.1, 5.0, ew - 0.2, 0.3, t, size=10)
    if i < 2:
        d.arrow(s, x + ew + 0.02, 5.01, x + ew + 0.18, 5.01, color=G)
d.notes(s, "判断：国内債券ややアンダー維持、先進国国債（ヘッジ有）ニュートラル維持。市場はFRBの年内1回＋2027年2回の利上げを織り込むが、CIOは年内1回・2027年なし。出典：p.2, p.4, p.14")

# 7 REIT ---------------------------------------------------------------
s = d.slide("REIT：金利上昇が逆風、先進国REITを引き下げ")
two_charts(s, "fig1_1.png", "fig1_2.png", "原資料 p.3 図表1-1", "原資料 p.3 図表1-2", h=3.3)
panel(s, "なぜ金利上昇に弱いのか", [
    "投資家がより高い分配金利回りを求め、価格が下がりやすい",
    "借入金の負担が増え、分配金の原資が減る",
    "先行するモーゲージREIT（右図）は足元やや軟調",
    "AI脅威論でデータセンター系REITにも逆風"], h=3.3)
d.callout(s, 0.4, 4.6, 9.2, 0.7, "判断", "先進国REIT：ややオーバー→ニュートラルに引き下げ。国内REITはややアンダー維持", color=N)
d.notes(s, "左図：長期金利（逆目盛）が上昇する局面で国内REITは株式に見劣り。出典：p.3, p.15")

# 8 AI・半導体 ----------------------------------------------------------
s = d.slide("AI・半導体：好況の裏で変調の兆し")
two_charts(s, "fig3_1.png", "fig3_2.png", "原資料 p.5 図表3-1", "原資料 p.5 図表3-2", h=3.3)
d.box(s, 6.35, 0.95, 3.25, 1.05, fill=C.RED_T)
d.text(s, 6.45, 0.98, 3.05, 0.28, "世界半導体売上高（前年比）", size=10, color=C.SUB, align="c")
d.text(s, 6.45, 1.25, 3.05, 0.5, [[("124.4%", {"color": C.SUB, "size": 16}), ("  →  ", {"color": R}),
                                    ("136.2%", {"color": R, "size": 22})]], size=16, bold=True, align="c", anchor="m")
d.text(s, 6.45, 1.72, 3.05, 0.25, "6月 → 7月（伸びがさらに加速）", size=9, color=C.SUB, align="c")
sig = [("出荷・在庫バランスがマイナス", "韓国の出荷が鈍り在庫が増加（左図）"),
       ("メモリ価格が頭打ち", "NANDは4月頃、DRAMも8月下旬から（右図）"),
       ("「AI脅威論」の台頭", "DC建設延期が続けば価格は一段安")]
for i, (h, b) in enumerate(sig):
    y = 2.12 + i * 0.73
    d.box(s, 6.35, y, 3.25, 0.66, fill=C.BG)
    d.num_badge(s, 6.43, y + 0.17, i + 1, color=N, d=0.32, size=11)
    d.text(s, 6.85, y + 0.03, 2.7, 0.3, h, size=10.5, bold=True, color=N, anchor="m")
    d.text(s, 6.85, y + 0.33, 2.7, 0.3, b, size=9.5, anchor="m")
d.callout(s, 0.4, 4.6, 9.2, 0.7, "ひとこと", "価格水準はまだ高く直ちに業績圧迫とは限らないが要警戒。新興国株式は「ややアンダー」維持", color=N)
d.notes(s, "出典：p.5 AI・半導体景気の変調に備える、p.12 新興国株式")

# 9 信用リスク ----------------------------------------------------------
s = d.slide("信用リスク：「悪循環」の入口に注意")
two_charts(s, "fig4_1.png", "fig4_2.png", "原資料 p.6 図表4-1", "原資料 p.6 図表4-2", h=3.2)
chk = [("米HYスプレッド", "低水準（未顕在化）", "平常", G), ("CCC格ローン上乗せ金利", "足元にかけて上昇", "注意", R),
       ("上場BDCの株価", "低水準で推移", "注意", R), ("BB・B格への波及", "波及なら判断を見直し", "監視", N)]
d.box(s, 6.35, 0.95, 3.25, 3.2, fill=C.BG)
d.text(s, 6.5, 1.0, 3.0, 0.32, "現状チェック", size=12, bold=True, color=N)
for i, (a, b, t, col) in enumerate(chk):
    y = 1.38 + i * 0.68
    d.tag(s, 6.48, y + 0.14, 0.58, t, color=col, h=0.27, size=9)
    d.text(s, 7.12, y, 2.4, 0.3, a, size=10.5, bold=True, anchor="m")
    d.text(s, 7.12, y + 0.29, 2.4, 0.28, b, size=9.5, color=C.SUB, anchor="m")
d.text(s, 0.4, 4.3, 5, 0.28, "最悪シナリオ：AI減速が信用不安に連鎖する悪循環", size=10.5, bold=True, color=C.SUB)
d.flow(s, 0.4, 4.62, 9.2, 0.7, ["AI投資の減速", "関連企業の業績悪化", "信用リスク上昇", "クラウド大手の収益圧迫",
                                 "設備投資の慎重化"], gap=0.25, size=10, colors=[N] * 5)
d.notes(s, "判断：ハイイールド社債ニュートラル維持、投資適格社債ややアンダー維持。出典：p.6, p.13")

# 10 原油・交易条件 ----------------------------------------------------
s = d.slide("原油高で交易条件に明暗、米国の優位が鮮明")
two_charts(s, "fig5_1.png", "fig5_2.png", "原資料 p.7 図表5-1", "原資料 p.7 図表5-2", h=3.3)
d.box(s, 6.35, 0.95, 3.25, 1.1, fill=C.BG)
d.text(s, 6.45, 0.98, 3.05, 0.28, "WTI原油（期近物）", size=10, color=C.SUB, align="c")
d.text(s, 6.45, 1.25, 3.05, 0.45, [[("約70", {"color": N}), ("  →  ", {"color": R}), ("106", {"color": R, "size": 22}),
                                    ("ドル台", {"color": R, "size": 12})]], size=18, bold=True, align="c", anchor="m")
d.text(s, 6.45, 1.72, 3.05, 0.28, "今年半ば → 9月に一時", size=9, color=C.SUB, align="c")
rows3 = [("米国", "▲ 改善", R, "原油輸出国。企業業績と株価に追い風"),
         ("日本", "▼ 悪化", N, "半年ほど遅れて粗利益率が低下"),
         ("中国", "▼ 大幅悪化", N, "日本以上に悪化し企業収益を圧迫")]
for i, (nm, t, col, b) in enumerate(rows3):
    y = 2.17 + i * 0.73
    d.box(s, 6.35, y, 3.25, 0.66, fill=C.RED_T if col == R else C.NAVY_T)
    d.text(s, 6.45, y + 0.03, 0.8, 0.3, nm, size=12, bold=True, color=col, anchor="m")
    d.text(s, 7.2, y + 0.03, 2.3, 0.3, t, size=10.5, bold=True, color=col, anchor="m")
    d.text(s, 6.45, y + 0.34, 3.1, 0.3, b, size=9.5, anchor="m")
d.callout(s, 0.4, 4.6, 9.2, 0.7, "判断", "米国の交易条件の優位性から、先進国株式は「ややオーバー」を維持", color=R)
d.notes(s, "交易条件＝輸出物価÷輸入物価。出典：p.7, p.12")

# 11 G>R ----------------------------------------------------------------
s = d.slide("日本株を支える「G>R」はむしろ拡大方向")
two_charts(s, "fig6_1.png", "fig6_2.png", "原資料 p.8 図表6-1", "原資料 p.8 図表6-2", h=3.3)
panel(s, "G>Rとは", [
    "名目成長率（G）が名目長期金利（R）を上回る状態",
    "過去、G>Rの局面では株式が国債より優位（左図）",
    "長期金利が3%近辺でもG>Rは維持",
    "マネーストックの伸びが加速し、G−Rの拡大を示唆（右図）"], h=3.3, color=R)
d.callout(s, 0.4, 4.6, 9.2, 0.7, "ひとこと", "銀行貸出→投資→所得→預金の好循環に期待。ただし日銀の利上げ加速が投資を冷やせば引き下げを検討", color=R)
d.notes(s, "出典：p.8 日本株は金利上昇でも「G>R」がサポート")

# 12 日本株 -------------------------------------------------------------
s = d.slide("日本株：業績は好調も、上値は緩やか")
d.image_fit(s, I("eps.png"), 0.4, 0.95, 5.6, 3.6, caption="原資料 p.11 国内株式の4局面とEPS")
d.box(s, 6.2, 0.95, 3.4, 1.85, fill=C.RED_T)
d.text(s, 6.32, 0.98, 3.2, 0.3, "▲ 追い風", size=12, bold=True, color=R)
d.text(s, 6.32, 1.3, 3.2, 1.5, ["決算良好、リビジョン指数は40%台（5年ぶり高水準）", "値上げの浸透", "TOPIXは8月に史上最高値"],
       size=10, bullet=True, space_after=3)
d.box(s, 6.2, 2.9, 3.4, 1.65, fill=C.NAVY_T)
d.text(s, 6.32, 2.93, 3.2, 0.3, "▼ 逆風", size=12, bold=True, color=N)
d.text(s, 6.32, 3.25, 3.2, 1.3, ["原油高による交易条件の悪化", "AI・半導体景気の変調", "日銀の利上げ加速"],
       size=10, bullet=True, space_after=3)
d.box(s, 2.6, 4.72, 4.8, 0.6, fill=G, text=[[("国内株式：ニュートラル維持", {"bold": True, "color": C.WHITE, "size": 15})]])
d.notes(s, "図はTOPIXの値動きを「流動性相場」「業績相場」「逆流動性相場」「逆業績相場」の4局面に分類したもの。出典：p.8, p.11")

# 13 債券 ---------------------------------------------------------------
s = d.slide("債券：スプレッドは割高、インフレ期待は2%超")
d.image_fit(s, I("spread.png"), 0.4, 0.95, 4.5, 2.6, caption="原資料 p.13 米国社債・新興国債券の利回りスプレッド")
d.image_fit(s, I("bei.png"), 5.1, 0.95, 4.5, 2.6, caption="原資料 p.14 日本の10年ブレークイーブンインフレ率")
bonds = [("投資適格社債", "ややアンダー", C.SCALE[1], "割高・金利上昇・大量発行"),
         ("ハイイールド社債", "ニュートラル", C.SCALE[2], "割高だが期間が短く高金利に強い"),
         ("新興国債券", "ややアンダー", C.SCALE[1], "割高・原油高・異常気象"),
         ("国内債券", "ややアンダー", C.SCALE[1], "日銀の利上げ・国債保有縮小・財政"),
         ("先進国国債（ヘッジ有）", "ニュートラル", C.SCALE[2], "市場は利上げを織り込み過ぎ")]
bw = (9.2 - 0.12 * 4) / 5
for i, (nm, v, col, why) in enumerate(bonds):
    x = 0.4 + i * (bw + 0.12)
    d.box(s, x, 3.75, bw, 1.58, fill=C.BG)
    d.text(s, x + 0.05, 3.8, bw - 0.1, 0.45, nm, size=10.5, bold=True, align="c", anchor="m")
    d.box(s, x + bw / 2 - 0.13, 4.27, 0.26, 0.26, fill=col, shape="oval")
    d.text(s, x + 0.05, 4.55, bw - 0.1, 0.25, v + "維持", size=9.5, align="c", color=C.SUB, bold=True)
    d.text(s, x + 0.08, 4.8, bw - 0.16, 0.5, why, size=9, align="c")
d.notes(s, "ブレークイーブンインフレ率＝市場が推測する将来のインフレ率。日銀の物価目標2%を上回って推移。出典：p.13, p.14")

# 14 1年間の推移 ---------------------------------------------------------
s = d.slide("ご参考：主要資産の1年間の値動き")
d.image_fit(s, I("p18.png"), 0.4, 0.95, 5.0, 4.4, caption="原資料 p.17（2025年8月＝100）")
d.box(s, 5.7, 0.95, 3.9, 2.45, fill=C.BG)
d.text(s, 5.85, 1.02, 3.6, 0.3, "図の見方", size=12, bold=True, color=N)
d.text(s, 5.85, 1.4, 3.6, 1.7, [
    "4つの資産クラスごとに、2025年8月を100として指数化した推移",
    "株式は各資産とも100を上回って推移",
    "金利資産の国内債券は100を下回る推移",
    "破線は為替ヘッジありの系列"], size=10.5, bullet=True, space_after=6)
d.text(s, 5.85, 3.0, 3.6, 0.3, "※ 傾向のみ記載（図から数値は読み取っていない）", size=9, color=C.GRAY)
d.notes(s, "出典：p.17 ご参考 1年間の主要資産の推移")

# 15 トリガー ------------------------------------------------------------
s = d.slide("今後の注目点：判断変更のトリガー")
rows = [["注目点（もし起きたら）", "想定される影響・判断"],
        ["AI・半導体の減速＋BB格・B格ローンの上乗せ金利が上昇", [[("ハイイールド社債の", {}), ("引き下げを検討", {"bold": True, "color": N})]]],
        ["日銀の利上げ加速で投資需要が冷え込む", [[("国内株式の", {}), ("引き下げを検討", {"bold": True, "color": N})]]],
        ["中東情勢が沈静化し原油価格が下落", "長期金利の上昇は一服（基調的な低下には景気後退が条件）"],
        ["日銀の国債保有縮小方針が変わらない", "日本の長期金利は相対的に上がりやすい地合いが継続"],
        ["円安是正に向けた政府・日銀の追加対応", "円高リスク（割高な資産の判断を順次引き下げ）"]]
d.table(s, 0.4, 0.95, 9.2, rows, [4.3, 4.9], row_h=0.5, size=10.5, header_size=11)
d.text(s, 0.4, 4.15, 9.2, 0.3, "日程が明らかな注目イベント", size=11, bold=True, color=C.SUB)
evs = [("2026年11月", "米中間選挙（トランプ氏はイランとの戦闘が選挙後まで続くと発言）"),
       ("2026年12月", "FRB・ECBの利上げ（CIO予想）"),
       ("2027年1月・4月", "日銀の利上げ（CIO予想）")]
ew = (9.2 - 0.2 * 2) / 3
for i, (w_, t) in enumerate(evs):
    x = 0.4 + i * (ew + 0.2)
    d.box(s, x, 4.5, ew, 0.8, fill=C.BG)
    d.text(s, x + 0.1, 4.52, ew - 0.2, 0.28, w_, size=10.5, bold=True, color=R)
    d.text(s, x + 0.1, 4.8, ew - 0.2, 0.48, t, size=9.5)
d.notes(s, "出典：p.4〜p.8, p.13〜p.14 の記述から、判断変更の条件として明記された箇所を抽出")

# 16 用語・出典 -----------------------------------------------------------
s = d.slide("用語ミニ解説と出典")
gl = [["用語", "意味"],
      ["G>R", "名目経済成長率(G)が名目長期金利(R)を上回る状態。株式が債券より優位になりやすい"],
      ["リビジョン・インデックス", "利益予想を上方修正した企業の割合－下方修正した企業の割合"],
      ["TAA", "戦術的資産配分。短期の見通しで配分を機動的に調整する"],
      ["モーゲージREIT", "不動産担保証券・ローンに投資するREIT。金利に敏感でREIT全体に先行"],
      ["BDC", "中堅・中小の未公開企業に投融資する投資会社"],
      ["交易条件", "輸出物価÷輸入物価。悪化は企業収益の重しになる"],
      ["ブレークイーブン\nインフレ率", "市場参加者が推測する将来のインフレ率"]]
d.table(s, 0.4, 0.95, 5.9, gl, [1.75, 4.15], row_h=0.5, size=9.5, header_size=10)
d.image_fill(s, I("photo_glass.jpg"), 6.55, 0.95, 3.05, 1.3)
d.box(s, 6.55, 2.25, 3.05, 3.1, fill=C.BG, shape="rect")
d.text(s, 6.7, 2.32, 2.8, 3.0, [
    [("出典", {"bold": True, "size": 11})],
    "野村證券「Nomura CIO Monthly」2026年10月号（投資判断は2026年9月18日時点）",
    "比較：同 2026年9月号（8月21日時点）",
    [("注記", {"bold": True, "size": 11})],
    "原資料の内容を個人的な理解のために要約。数値・判断・図表・写真は原資料による",
    "地域のトーン分類と図解は要約者による整理"], size=9.5, space_after=4)
d.save(OUT)
print("saved", OUT)
