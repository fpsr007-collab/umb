from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── ページ設定 ────────────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Cm(21.0)
section.page_height = Cm(29.7)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.0)

# ── フォント・スタイル共通設定 ──────────────────────────────
FONT_NAME    = "游明朝"
FONT_NAME_EN = "Times New Roman"
COLOR_NAVY   = RGBColor(0x1F, 0x38, 0x64)   # 濃紺
COLOR_ACCENT = RGBColor(0xC0, 0x00, 0x00)   # 深紅
COLOR_GRAY   = RGBColor(0x59, 0x59, 0x59)   # ダークグレー
COLOR_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_LIGHTBLUE = RGBColor(0xDC, 0xE6, 0xF1)
COLOR_LIGHTYELLOW = RGBColor(0xFF, 0xF2, 0xCC)
COLOR_LIGHTRED  = RGBColor(0xFF, 0xE0, 0xE0)
COLOR_LIGHTGREEN = RGBColor(0xE2, 0xEF, 0xDA)


def set_font(run, size, bold=False, color=None, name=None):
    run.font.name = name or FONT_NAME
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), name or FONT_NAME)
    rPr.insert(0, rFonts)


def para_space(para, before=0, after=0, line=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line:
        pf.line_spacing = Pt(line)


def shade_cell(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        if val:
            el = OxmlElement(f"w:{side}")
            for k, v in val.items():
                el.set(qn(f"w:{k}"), v)
            tcBorders.append(el)
    tcPr.append(tcBorders)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    para_space(p, before=14, after=6)
    if level == 1:
        run = p.add_run(text)
        set_font(run, 16, bold=True, color=COLOR_NAVY)
        p.paragraph_format.left_indent = Cm(0)
        # 下線装飾
        border = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"),   "single")
        bottom.set(qn("w:sz"),    "6")
        bottom.set(qn("w:space"), "4")
        bottom.set(qn("w:color"), "1F3864")
        border.append(bottom)
        p._p.get_or_add_pPr().append(border)
    elif level == 2:
        run = p.add_run(f"■ {text}")
        set_font(run, 12, bold=True, color=COLOR_NAVY)
        p.paragraph_format.left_indent = Cm(0.3)
        para_space(p, before=10, after=4)
    elif level == 3:
        run = p.add_run(f"◆ {text}")
        set_font(run, 11, bold=True, color=COLOR_ACCENT)
        p.paragraph_format.left_indent = Cm(0.5)
        para_space(p, before=8, after=3)
    return p


def add_body(doc, text, indent=0.5, size=10.5, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    para_space(p, before=2, after=2, line=16)
    run = p.add_run(text)
    set_font(run, size, color=color or COLOR_GRAY)
    return p


def add_bullet(doc, text, indent=0.8, marker="・"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(indent)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    para_space(p, before=1, after=1, line=15)
    run = p.add_run(f"{marker}　{text}")
    set_font(run, 10.5, color=COLOR_GRAY)
    return p


# ════════════════════════════════════════════════════════════
#  表紙
# ════════════════════════════════════════════════════════════

# タイトルブロック（表で装飾）
tbl = doc.add_table(rows=1, cols=1)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.style = "Table Grid"
cell = tbl.cell(0, 0)
shade_cell(cell, "1F3864")
cell.width = Cm(16)

cp = cell.paragraphs[0]
cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_space(cp, before=20, after=6)
run = cp.add_run("就　業　規　則　整　備　の　ご　提　案")
set_font(run, 20, bold=True, color=COLOR_WHITE)

cp2 = cell.add_paragraph()
cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
para_space(cp2, before=4, after=20)
run2 = cp2.add_run("〜 従業員10人未満の企業における就業規則未整備のリスクと対策 〜")
set_font(run2, 11, color=RGBColor(0xBD, 0xD7, 0xEE))

doc.add_paragraph()

# 提案者・日付情報
info_tbl = doc.add_table(rows=4, cols=2)
info_tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT
widths = [Cm(3.5), Cm(8)]
for row in info_tbl.rows:
    for i, cell in enumerate(row.cells):
        cell.width = widths[i]

info_items = [
    ("提　案　日", "令和　　年　　月　　日"),
    ("提　案　者", ""),
    ("宛　　　先", ""),
    ("作　成　者", ""),
]
for ri, (label, val) in enumerate(info_items):
    lc = info_tbl.cell(ri, 0)
    vc = info_tbl.cell(ri, 1)
    shade_cell(lc, "DCE6F1")
    lp = lc.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lr = lp.add_run(label)
    set_font(lr, 10, bold=True, color=COLOR_NAVY)
    vp = vc.paragraphs[0]
    vr = vp.add_run(val)
    set_font(vr, 10, color=COLOR_GRAY)

doc.add_paragraph()

# 概要ボックス
summary_tbl = doc.add_table(rows=1, cols=1)
summary_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
sc = summary_tbl.cell(0, 0)
shade_cell(sc, "FFF2CC")
sp = sc.paragraphs[0]
para_space(sp, before=8, after=2)
sr = sp.add_run("【本提案書の目的】")
set_font(sr, 11, bold=True, color=COLOR_ACCENT)
sp2 = sc.add_paragraph()
para_space(sp2, before=2, after=8)
sr2 = sp2.add_run(
    "従業員が10人未満の企業には、労働基準法上の就業規則作成義務はありません。しかし、"
    "就業規則を整備していないことにより、労使トラブル・助成金不受給・採用力低下など、"
    "企業経営に深刻な影響を与えるリスクが生じます。本提案書では、そのリスクを具体的に"
    "示すとともに、早期の就業規則整備をご提案いたします。"
)
set_font(sr2, 10.5, color=COLOR_GRAY)
sp2.paragraph_format.left_indent = Cm(0.3)

doc.add_page_break()

# ════════════════════════════════════════════════════════════
#  第１章　法的背景
# ════════════════════════════════════════════════════════════
add_heading(doc, "第１章　法的背景と義務の範囲", level=1)

add_body(doc,
    "労働基準法第89条は、常時10人以上の従業員を使用する事業場に対し、就業規則の作成および"
    "所轄労働基準監督署への届出を義務付けています。この規定により、従業員が10人未満の企業は"
    "法的な作成義務を負いません。",
    indent=0.3)

add_body(doc,
    "しかしながら、「義務がない＝整備しなくてよい」ではありません。就業規則は、会社と従業員"
    "双方のルールブックであり、その存在が「労使間の紛争解決の基準」「懲戒処分の根拠」"
    "「各種制度導入の法的要件」として機能します。",
    indent=0.3)

# 法的根拠表
add_heading(doc, "関連法規の概要", level=2)
law_tbl = doc.add_table(rows=5, cols=2)
law_tbl.style = "Table Grid"
law_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
col_w = [Cm(5), Cm(11)]
headers = ["法律・条文", "内容"]
shade_cell(law_tbl.cell(0, 0), "1F3864")
shade_cell(law_tbl.cell(0, 1), "1F3864")
for ci, h in enumerate(headers):
    hp = law_tbl.cell(0, ci).paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run(h)
    set_font(hr, 10, bold=True, color=COLOR_WHITE)

law_data = [
    ("労基法 第89条", "常時10人以上の事業場に就業規則の作成・届出義務"),
    ("労基法 第106条", "就業規則の労働者への周知義務（10人未満も適用）"),
    ("労働契約法 第7条", "就業規則が労働契約の内容となる条件を規定"),
    ("労働契約法 第10条", "就業規則変更による不利益変更の要件を規定"),
]
for ri, (law, desc) in enumerate(law_data, start=1):
    shade_cell(law_tbl.cell(ri, 0), "DCE6F1")
    lp = law_tbl.cell(ri, 0).paragraphs[0]
    lr = lp.add_run(law)
    set_font(lr, 10, bold=True, color=COLOR_NAVY)
    dp = law_tbl.cell(ri, 1).paragraphs[0]
    dr = dp.add_run(desc)
    set_font(dr, 10, color=COLOR_GRAY)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  第２章　リスク詳細
# ════════════════════════════════════════════════════════════
add_heading(doc, "第２章　就業規則未整備による主なリスク", level=1)

risks = [
    {
        "title": "リスク① 懲戒処分が法的に無効となるリスク",
        "severity": "★★★（深刻）",
        "body": (
            "裁判所・労働審判の判例（フジ興産事件・最高裁2003年）では、「懲戒処分は就業規則等に"
            "懲戒事由と処分の種類が明記されていることが有効要件」とされています。"
        ),
        "bullets": [
            "問題社員への減給・出勤停止・懲戒解雇が就業規則なしでは無効と判断される可能性が高い",
            "解雇無効と判断された場合、解雇期間中の賃金（バックペイ）を全額支払う義務が生じる",
            "訴訟費用・弁護士費用を含めた経営的ダメージは中小企業にとって致命的となり得る",
        ],
        "fill": "FFE0E0",
    },
    {
        "title": "リスク② 解雇・雇止めトラブルへの対応力低下",
        "severity": "★★★（深刻）",
        "body": (
            "解雇の理由・手続き・予告期間などのルールが書面で存在しないため、「不当解雇」を"
            "主張された際に会社側の法的根拠が著しく弱くなります。"
        ),
        "bullets": [
            "有期契約の更新・雇止めの基準が不明確で、労働審判・訴訟で不利な立場になりやすい",
            "口頭のみの約束では「言った・言わない」の水掛け論となり、証拠能力がない",
            "労働基準監督署の調査時に書面根拠がなく、是正勧告・指導票を受けるリスクがある",
        ],
        "fill": "FFE0E0",
    },
    {
        "title": "リスク③ 固定残業代（みなし残業）の無効リスク",
        "severity": "★★★（深刻）",
        "body": (
            "固定残業代を給与に含める場合、その計算根拠・対象時間数・超過時の追加払いルールが"
            "就業規則または労働条件通知書に明記されていることが有効要件のひとつです。"
        ),
        "bullets": [
            "就業規則がなく記載も不十分な場合、固定残業代が全額「基本給」と見なされる",
            "過去2〜3年分の未払い残業代＋遅延損害金の遡及請求を受けるリスクがある",
            "数年分の残業代請求は、従業員数名でも数百万〜数千万円規模になり得る",
        ],
        "fill": "FFE0E0",
    },
    {
        "title": "リスク④ 変形労働時間制・フレックスタイム制を導入できない",
        "severity": "★★☆（中程度）",
        "body": (
            "シフト制や繁閑対応の柔軟な勤務体系を導入するには、就業規則への記載が"
            "労働基準法上の要件となっています（第32条の2・32条の4等）。"
        ),
        "bullets": [
            "就業規則なしに変形労働時間制を運用しても法的効力がなく、全時間に残業代が発生する",
            "季節・繁閑に合わせた人員配置の最適化ができず、余分な人件費が発生する",
        ],
        "fill": "FFF2CC",
    },
    {
        "title": "リスク⑤ 雇用関係助成金を受給できない",
        "severity": "★★☆（中程度）",
        "body": (
            "厚生労働省の多くの雇用関係助成金では、就業規則の整備・届出が申請要件に"
            "含まれています。整備していないだけで申請資格を失います。"
        ),
        "bullets": [
            "キャリアアップ助成金（正社員転換コース）：非正規→正規転換で最大80万円/人",
            "両立支援等助成金：育児・介護休業制度整備で最大100万円",
            "人材確保等支援助成金：処遇改善で最大570万円",
            "就業規則がないだけで、これらの受給機会を逃すことになる",
        ],
        "fill": "FFF2CC",
    },
    {
        "title": "リスク⑥ 採用力・人材定着率の低下",
        "severity": "★★☆（中程度）",
        "body": (
            "転職経験のある求職者は就業規則の開示を求めることがあります。開示できない場合、"
            "職場環境が不透明と判断され、優秀な人材の採用に影響します。"
        ),
        "bullets": [
            "「ブラック企業」と誤解されるリスクがあり、採用競争力が低下する",
            "入社後に労働条件のミスマッチが発覚し、早期離職・採用コストの無駄が増える",
            "従業員が自分の権利を確認できず、不安・不満から離職率が上がりやすい",
        ],
        "fill": "FFF2CC",
    },
    {
        "title": "リスク⑦ 10人到達時に即日義務発生・未届出は罰則あり",
        "severity": "★★☆（中程度）",
        "body": (
            "従業員が10人以上となった時点で、就業規則の作成・届出義務が即日発生します。"
            "未届出のまま放置した場合、労働基準法第120条により罰則が適用されます。"
        ),
        "bullets": [
            "違反した場合：30万円以下の罰金（労基法第120条）",
            "事前に整備しておけば、10人到達後に慌てて作成するコスト・リスクを避けられる",
            "従業員数の変動が多い成長企業ほど、早期整備が重要",
        ],
        "fill": "FFF2CC",
    },
]

for risk in risks:
    add_heading(doc, risk["title"], level=3)
    # 深刻度バッジ
    sp = doc.add_paragraph()
    sp.paragraph_format.left_indent = Cm(0.8)
    para_space(sp, before=0, after=4)
    sr = sp.add_run(f"深刻度：{risk['severity']}")
    set_font(sr, 10, bold=True, color=COLOR_ACCENT)

    add_body(doc, risk["body"], indent=0.8)
    for b in risk["bullets"]:
        add_bullet(doc, b, indent=1.0)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  第３章　リスク一覧表
# ════════════════════════════════════════════════════════════
add_heading(doc, "第３章　リスク一覧サマリー", level=1)

summary_rows = [
    ("懲戒処分の無効",         "問題社員への対処不能・バックペイ請求",   "★★★"),
    ("解雇トラブル",           "不当解雇訴訟・バックペイ・復職命令",     "★★★"),
    ("固定残業代の無効",       "数年分の未払い残業代の遡及請求",         "★★★"),
    ("変形労働時間制の不可",   "余分な残業代の発生・人件費増",           "★★☆"),
    ("助成金の不受給",         "数十万〜数百万円の機会損失",             "★★☆"),
    ("採用力・定着率の低下",   "優秀な人材の確保困難・早期離職増加",     "★★☆"),
    ("10人到達時の未届出罰則", "30万円以下の罰金",                       "★★☆"),
]

tbl = doc.add_table(rows=len(summary_rows)+1, cols=3)
tbl.style = "Table Grid"
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
col_ws = [Cm(4.5), Cm(8), Cm(3)]

for ci, hdr in enumerate(["リスク項目", "具体的な影響", "深刻度"]):
    shade_cell(tbl.cell(0, ci), "1F3864")
    hp = tbl.cell(0, ci).paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run(hdr)
    set_font(hr, 10, bold=True, color=COLOR_WHITE)

for ri, (item, impact, severity) in enumerate(summary_rows, start=1):
    fill = "FFE0E0" if "★★★" in severity else "FFF9E6"
    for ci in range(3):
        shade_cell(tbl.cell(ri, ci), fill)
    ip = tbl.cell(ri, 0).paragraphs[0]
    ir = ip.add_run(item)
    set_font(ir, 10, bold=True, color=COLOR_NAVY)
    ep = tbl.cell(ri, 1).paragraphs[0]
    er = ep.add_run(impact)
    set_font(er, 10, color=COLOR_GRAY)
    sp2 = tbl.cell(ri, 2).paragraphs[0]
    sp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr2 = sp2.add_run(severity)
    set_font(sr2, 10, bold=True, color=COLOR_ACCENT)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  第４章　提案・対策
# ════════════════════════════════════════════════════════════
add_heading(doc, "第４章　ご提案：就業規則の早期整備", level=1)

add_body(doc,
    "以上のリスクを踏まえ、従業員10人未満の段階から就業規則を整備することを強くお勧めします。"
    "以下に整備にあたっての推奨事項をご提案いたします。",
    indent=0.3)

add_heading(doc, "最低限整備すべき項目", level=2)
must_items = [
    "労働時間・休憩・休日・休暇（年次有給休暇を含む）",
    "賃金の決定・計算・支払方法・締切日・支払日",
    "退職・解雇の事由と手続き",
    "懲戒の種類・事由・手続き",
    "試用期間・本採用の基準",
    "ハラスメント防止・相談窓口",
    "秘密保持・個人情報の取り扱い",
]
for item in must_items:
    add_bullet(doc, item, marker="✓")

add_heading(doc, "整備スケジュール（例）", level=2)
steps = [
    ("STEP 1（〜1か月目）", "現状の雇用形態・賃金体系・労働時間を棚卸しし、整備すべき事項を洗い出す"),
    ("STEP 2（〜2か月目）", "社会保険労務士・弁護士と就業規則の草案を作成する"),
    ("STEP 3（〜3か月目）", "従業員代表の意見聴取を行い、内容を確定する"),
    ("STEP 4（〜4か月目）", "就業規則を従業員に周知し、運用を開始する"),
    ("STEP 5（随時）",       "法改正・会社状況の変化に応じて定期的に見直す"),
]
for step, desc in steps:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    para_space(p, before=3, after=3, line=15)
    r1 = p.add_run(f"{step}　")
    set_font(r1, 10, bold=True, color=COLOR_NAVY)
    r2 = p.add_run(desc)
    set_font(r2, 10, color=COLOR_GRAY)

add_heading(doc, "整備コストと費用対効果", level=2)
cost_tbl = doc.add_table(rows=4, cols=3)
cost_tbl.style = "Table Grid"
for ci, hdr in enumerate(["整備方法", "概算費用", "メリット"]):
    shade_cell(cost_tbl.cell(0, ci), "1F3864")
    hp = cost_tbl.cell(0, ci).paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hr = hp.add_run(hdr)
    set_font(hr, 10, bold=True, color=COLOR_WHITE)
cost_data = [
    ("自社作成（テンプレート活用）", "数万円程度（印刷・労力）", "コスト最小・カスタマイズ自由"),
    ("社会保険労務士に依頼",         "10〜30万円程度",         "法的整合性が高く・アドバイス付き"),
    ("弁護士に依頼",                 "30〜100万円程度",         "訴訟リスク対策まで含めた高品質"),
]
fills = ["E2EFDA", "FFF2CC", "DCE6F1"]
for ri, (method, cost, merit) in enumerate(cost_data, start=1):
    shade_cell(cost_tbl.cell(ri, 0), fills[ri-1])
    shade_cell(cost_tbl.cell(ri, 1), fills[ri-1])
    shade_cell(cost_tbl.cell(ri, 2), fills[ri-1])
    for ci, val in enumerate([method, cost, merit]):
        p = cost_tbl.cell(ri, ci).paragraphs[0]
        r = p.add_run(val)
        set_font(r, 10, color=COLOR_GRAY)

doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  結語
# ════════════════════════════════════════════════════════════
add_heading(doc, "おわりに", level=1)

conc_tbl = doc.add_table(rows=1, cols=1)
conc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
cc = conc_tbl.cell(0, 0)
shade_cell(cc, "DCE6F1")
cp1 = cc.paragraphs[0]
para_space(cp1, before=10, after=4)
cr1 = cp1.add_run("就業規則の整備は、企業を守る「盾」であり、従業員が安心して働ける「基盤」です。")
set_font(cr1, 11, bold=True, color=COLOR_NAVY)
cp1.alignment = WD_ALIGN_PARAGRAPH.CENTER

cp2 = cc.add_paragraph()
para_space(cp2, before=4, after=10)
cr2 = cp2.add_run(
    "整備コストは数万〜数十万円であるのに対し、未整備によるリスク（残業代訴訟・解雇無効・"
    "助成金不受給）はその何倍もの損失につながります。従業員が10人に達した時点で届出義務が"
    "即日発生することからも、早期の整備が最も合理的な経営判断です。\n"
    "ご不明な点がございましたら、お気軽にご相談ください。"
)
set_font(cr2, 10.5, color=COLOR_GRAY)
cp2.paragraph_format.left_indent = Cm(0.5)

# ── 保存 ─────────────────────────────────────────────────────
output = "/home/user/umb/就業規則整備の提案書.docx"
doc.save(output)
print(f"Saved: {output}")
