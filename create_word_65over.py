from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3)
section.right_margin = Cm(3)

style = doc.styles['Normal']
style.font.name = '游明朝'
style.font.size = Pt(10.5)

def set_font(run, bold=False, size=None, color=None):
    run.font.name = '游明朝'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')
    if bold:
        run.bold = True
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, bold=True, size=14, color=(0x1F, 0x49, 0x7D))
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_body(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_bullet(doc, parts):
    p = doc.add_paragraph(style='List Bullet')
    for text, bold in parts:
        run = p.add_run(text)
        set_font(run, bold=bold)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_box(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = Pt(18)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'DEEAF1')
    pPr.append(shd)
    return p

def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for para in hdr_cells[i].paragraphs:
            for run in para.runs:
                set_font(run, bold=True)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tc = hdr_cells[i]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '2E74B5')
        tcPr.append(shd)
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = val
            for para in row_cells[c_idx].paragraphs:
                for run in para.runs:
                    set_font(run)
        if r_idx % 2 == 1:
            for cell in row_cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), 'DEEAF1')
                tcPr.append(shd)
    return table

# ===== 本文 =====

# タイトル
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = p_title.add_run(
    '70歳までの就業機会確保を検討している事業主へ\n65歳超雇用推進助成金を活用しましょう'
)
set_font(run_title, bold=True, size=15, color=(0x1F, 0x49, 0x7D))
p_title.paragraph_format.space_after = Pt(10)

# リード
add_body(doc,
    '厚生労働省は、2026年度の重点課題として「多様な人材の活躍促進」を掲げており、'
    '70歳までの就業機会確保や高齢者の処遇改善に取り組む企業への支援を強化しています。'
    '少子高齢化が進む中、高齢者の豊富な経験・知識を活かした戦力化は、人手不足対策としても重要な経営課題です。'
    '「65歳超雇用推進助成金」を活用して、高齢者が長く活躍できる職場づくりを進めましょう。'
)

# 見出し1
add_heading(doc, '◎ なぜ今、70歳までの就業機会確保が求められるのか')

add_body(doc,
    '2021年4月に改正高年齢者雇用安定法が施行され、70歳までの就業機会確保が企業の「努力義務」となりました。'
    '「義務」ではないものの、今後の法改正の動向や人材確保の観点から、早めに対応を検討しておくことが重要です。'
)

add_bullet(doc, [('65歳までの雇用確保', True), ('：継続雇用制度の導入または定年の引き上げが「義務」', False)])
add_bullet(doc, [('70歳までの就業機会確保', True), ('：定年延長・継続雇用・業務委託・社会貢献活動参加等が「努力義務」', False)])

add_box(doc,
    '【背景】\n'
    '日本の総人口に占める65歳以上の割合は約3割に達しており、高齢者の就労促進は社会保障の持続性にとっても重要なテーマです。'
    '政府は「生涯現役社会」の実現に向け、企業の取り組みを助成金で後押ししています。'
)

# 見出し2
add_heading(doc, '◎ 65歳超雇用推進助成金とは')

add_body(doc,
    '65歳以上への定年引き上げや高年齢者の雇用環境整備、処遇改善に取り組む事業主を支援する助成金です。'
    '以下の３つのコースで構成されており、自社の状況に応じて活用できます。'
)

add_table(doc,
    ['コース名', '主な対象となる取り組み'],
    [
        ['65歳超継続雇用促進コース',
         '65歳以上への定年引き上げ、定年の廃止、希望者全員を対象とする66歳以上の継続雇用制度の導入など'],
        ['高年齢者評価制度等雇用管理改善コース',
         '高年齢者の雇用管理制度の整備（能力・職務評価制度の導入、健康・体力測定の実施等）'],
        ['高年齢者無期雇用転換コース',
         '50歳以上かつ定年年齢未満の有期契約労働者を無期雇用労働者に転換'],
    ]
)

doc.add_paragraph()

# 見出し3
add_heading(doc, '◎ 各コースの支給額の目安')

add_table(doc,
    ['コース', '支給額の目安'],
    [
        ['65歳超継続雇用促進コース', '定年引き上げ幅・対象人数に応じて10万〜160万円'],
        ['高年齢者評価制度等雇用管理改善コース', '支給対象経費の60％（中小企業）、最大60万円'],
        ['高年齢者無期雇用転換コース', '対象労働者１人あたり最大30万円'],
    ]
)

doc.add_paragraph()

add_body(doc,
    '※支給額・要件は変更される場合があります。最新情報は厚生労働省または都道府県労働局にご確認ください。'
)

# 見出し4
add_heading(doc, '◎ 高齢者の処遇改善も重要なポイント')

add_body(doc,
    '高齢者を長く戦力として活躍してもらうためには、雇用継続の制度を整えるだけでなく、処遇面の改善も欠かせません。'
    '定年後の賃金が大幅に下がるいわゆる「定年後再雇用による処遇低下」は、モチベーション低下や優秀な人材の流出につながる恐れがあります。'
)

add_bullet(doc, [('同一労働同一賃金の観点', True), ('から、職務内容に見合った処遇の設定が求められます', False)])
add_bullet(doc, [('能力・成果に基づく評価制度', True), ('を導入し、年齢に関わらず意欲を持って働ける環境を整備しましょう', False)])
add_bullet(doc, [('健康管理・職場環境の整備', True), ('も、高齢者が安心して長く働くために重要な取り組みです', False)])

# 見出し5
add_heading(doc, '◎ 活用の流れ（65歳超継続雇用促進コースの場合）')

add_bullet(doc, [('STEP 1：', True), ('自社の定年・継続雇用制度の現状を確認し、引き上げ・改定の方針を検討', False)])
add_bullet(doc, [('STEP 2：', True), ('就業規則を改定し、労働者への周知を実施', False)])
add_bullet(doc, [('STEP 3：', True), ('就業規則の改定日から２か月以内に支給申請書を都道府県労働局へ提出', False)])
add_bullet(doc, [('STEP 4：', True), ('審査・支給決定', False)])

add_box(doc,
    '【実務上の注意点】\n'
    '・就業規則の改定前に申請手続きの要件を必ず確認してください。\n'
    '・申請期限（改定日から２か月以内）を過ぎると支給対象外となります。\n'
    '・コースによって申請のタイミングや添付書類が異なります。詳細は最寄りの都道府県労働局または社会保険労務士にご相談ください。'
)

# まとめ
add_heading(doc, '◎ まとめ')

add_body(doc,
    '高齢者の雇用延長・処遇改善は、人手不足の解消や職場の活性化につながる重要な取り組みです。'
    '「65歳超雇用推進助成金」は、制度整備のコスト負担を軽減しながら、高齢者が長く安心して働ける職場環境を実現するための有効な手段です。'
    '70歳就業時代に備え、まずは自社の現行制度の確認と、活用できるコースの検討から始めてみましょう。'
)

output_path = '/home/user/umb/65歳超雇用推進助成金_解説.docx'
doc.save(output_path)
print(f'saved: {output_path}')
