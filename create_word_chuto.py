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

def add_indent(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run(text)
    set_font(run)
    p.paragraph_format.space_after = Pt(5)
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

def add_table(doc, headers, rows):
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
    '中途採用を拡大したい事業主へ\n早期再就職支援等助成金（中途採用拡大コース）を活用しましょう'
)
set_font(run_title, bold=True, size=15, color=(0x1F, 0x49, 0x7D))
p_title.paragraph_format.space_after = Pt(10)

# リード
add_body(doc,
    '厚生労働省は、2026年度の重点課題への対応方針として、賃金上昇を伴う中途採用者の雇用拡大を図る事業主への支援を掲げています。'
    'その中心となるのが「早期再就職支援等助成金（中途採用拡大コース）」です。'
    '人材確保に課題を抱える企業にとって、積極的に活用したい制度です。'
)

# 見出し1
add_heading(doc, '◎ 早期再就職支援等助成金（中途採用拡大コース）とは')

add_body(doc,
    '離職者の早期再就職を促進するとともに、企業における中途採用を拡大することを目的とした助成金です。'
    '中途採用者の採用比率を拡大する計画を作成・実施し、'
    '採用した中途採用者の賃金を一定水準以上に設定した場合に支給されます。'
)

# 見出し2
add_heading(doc, '◎ 支給の主な要件')

add_bullet(doc, [('中途採用拡大計画', True), ('を作成し、都道府県労働局長の認定を受けること', False)])
add_bullet(doc, [('計画期間中に中途採用者の採用比率を拡大', True), ('すること（直近３年間の平均採用比率を上回ること）', False)])
add_bullet(doc, [('採用した中途採用者の賃金', True), ('が、同種の業務に従事する既存労働者の賃金水準以上であること', False)])
add_bullet(doc, [('雇用保険適用事業主', True), ('であること', False)])

add_box(doc,
    '【ポイント】\n'
    '計画の認定を受ける前に採用した中途採用者は助成対象外となります。必ず採用活動の前に計画を提出・認定を受けてください。'
)

# 見出し3
add_heading(doc, '◎ 支給額の目安')

add_table(doc,
    ['区分', '支給額（1人あたり）'],
    [
        ['中途採用拡大コース（基本）', '最大50万円'],
        ['45歳以上の中途採用者を雇い入れた場合', '最大60万円（加算あり）'],
        ['生産性要件を満たす場合', '上記の3/4相当を上乗せ加算'],
    ]
)

doc.add_paragraph()

add_body(doc,
    '※支給額・要件は変更される場合があります。最新情報は厚生労働省または都道府県労働局にご確認ください。'
)

# 見出し4
add_heading(doc, '◎ 「賃金上昇を伴う」採用がポイント')

add_body(doc,
    '本コースの特徴は、単なる中途採用の拡大にとどまらず、「賃金上昇を伴う」ことを要件としている点です。'
    '昨今の人手不足・賃上げの流れを背景に、政府は中途採用市場の活性化と、転職者が処遇改善を実感できる環境づくりを一体で推進しています。'
)

add_bullet(doc, [('自社の賃金水準を見直すきっかけ', True), ('としても活用できる', False)])
add_bullet(doc, [('採用コストの一部を助成金で補填', True), ('しながら、優秀な中途人材を確保できる', False)])
add_bullet(doc, [('既存従業員の賃金とのバランス調整', True), ('も含めた人事制度の見直しを検討する好機', False)])

# 見出し5
add_heading(doc, '◎ 活用の流れ')

add_bullet(doc, [('STEP 1：', True), ('中途採用拡大計画を作成し、都道府県労働局へ提出・認定申請', False)])
add_bullet(doc, [('STEP 2：', True), ('認定後、計画に基づき中途採用活動を実施', False)])
add_bullet(doc, [('STEP 3：', True), ('採用した中途採用者の賃金が要件を満たすことを確認', False)])
add_bullet(doc, [('STEP 4：', True), ('計画期間終了後、支給申請書を提出', False)])

add_box(doc,
    '【実務上の注意点】\n'
    '・申請前に就業規則・賃金規定が整備されていることが前提となります。\n'
    '・採用した労働者が一定期間在籍していることが支給要件となる場合があります。\n'
    '・詳細は最寄りの都道府県労働局またはハローワーク、社会保険労務士にご相談ください。'
)

# まとめ
add_heading(doc, '◎ まとめ')

add_body(doc,
    '中途採用の拡大を検討している事業主にとって、「早期再就職支援等助成金（中途採用拡大コース）」は採用コストの負担軽減と人材確保を同時に実現できる有力な制度です。'
    '賃金水準の引き上げと合わせて取り組むことで、採用競争力の向上にもつながります。'
    '計画認定の申請手順は手続きが多いため、早めに専門家や労働局へ相談することをおすすめします。'
)

output_path = '/home/user/umb/早期再就職支援等助成金_中途採用拡大コース.docx'
doc.save(output_path)
print(f'saved: {output_path}')
