from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ページ余白設定
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3)
section.right_margin = Cm(3)

# デフォルトフォント設定
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

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    if level == 1:
        set_font(run, bold=True, size=14, color=(0x1F, 0x49, 0x7D))
    elif level == 2:
        set_font(run, bold=True, size=11.5, color=(0x2E, 0x74, 0xB5))
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_body(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run, bold=bold)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_bullet(doc, text, bold_part=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_part and bold_part in text:
        idx = text.index(bold_part)
        before = text[:idx]
        after = text[idx+len(bold_part):]
        if before:
            r1 = p.add_run(before)
            set_font(r1)
        r2 = p.add_run(bold_part)
        set_font(r2, bold=True)
        if after:
            r3 = p.add_run(after)
            set_font(r3)
    else:
        r = p.add_run(text)
        set_font(r)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_table(doc, headers, rows):
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'

    # ヘッダー行
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for para in hdr_cells[i].paragraphs:
            for run in para.runs:
                set_font(run, bold=True)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 背景色
        tc = hdr_cells[i]._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '2E74B5')
        tcPr.append(shd)
        for para in hdr_cells[i].paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # データ行
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx+1].cells
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = val
            for para in row_cells[c_idx].paragraphs:
                for run in para.runs:
                    set_font(run)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 偶数行に薄い背景
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

# ===== 本文開始 =====

# タイトル
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = p_title.add_run('社会保険の適用拡大、自社の対応スケジュールは確認できていますか？')
set_font(run_title, bold=True, size=16, color=(0x1F, 0x49, 0x7D))
p_title.paragraph_format.space_after = Pt(8)

# リード文
add_body(doc,
    '社会保険（健康保険・厚生年金保険）の適用拡大が、段階的に進められています。これまで「従業員51人以上」の企業が対象でしたが、今後は中小・零細企業にも順次適用が広がります。自社がいつ対象になるかを早めに把握し、準備を進めることが重要です。'
)

# 見出し1
add_heading(doc, '◎ 対象企業の拡大スケジュール', level=1)

add_body(doc, '適用対象となる企業規模（従業員数）は、以下のとおり段階的に引き下げられます。')

# テーブル
add_table(doc,
    ['従業員数', '施行時期'],
    [
        ['51人以上', '適用済み'],
        ['36〜50人', '令和９年（2027年）10月〜'],
        ['21〜35人', '令和11年（2029年）10月〜'],
        ['11〜20人', '令和14年（2032年）10月〜'],
        ['10人以下', '令和17年（2035年）10月〜'],
    ]
)

doc.add_paragraph()  # スペース

add_body(doc,
    '「まだ先の話」と感じる方もいるかもしれませんが、社会保険料の負担は企業・従業員の双方に影響します。採用・労務管理の見直しが必要になる場合もあるため、早めの試算と準備をおすすめします。'
)

# 見出し2
add_heading(doc, '◎ 「賃金要件」の撤廃にも要注意', level=1)

add_body(doc,
    '適用対象となる従業員に関する要件についても、重要な変更があります。'
)

p = doc.add_paragraph()
r1 = p.add_run('令和８年（2026年）10月以降、これまで加入要件の一つであった')
set_font(r1)
r2 = p.add_run('「所定内賃金が月額8.8万円以上」という要件が撤廃')
set_font(r2, bold=True)
r3 = p.add_run('されます。')
set_font(r3)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = Pt(18)

add_body(doc, '以下の２つの要件を満たすパート・アルバイト等の短時間労働者は、賃金額にかかわらず社会保険の加入対象となります。')

add_bullet(doc, '週所定労働時間が20時間以上であること', bold_part='週所定労働時間が20時間以上')
add_bullet(doc, '学生でないこと', bold_part='学生でない')

add_body(doc,
    'たとえば、週20時間以上働いているものの月収が8.8万円未満のスタッフも、令和８年10月以降は加入対象となる可能性があります。現在のシフト・雇用条件を今一度確認しておきましょう。'
)

# 見出し3
add_heading(doc, '◎ 企業側の主な影響と対応ポイント', level=1)

add_body(doc, '社会保険の適用拡大に伴い、企業には以下のような影響が生じる可能性があります。')

add_bullet(doc, '保険料の事業主負担が増加する', bold_part='保険料の事業主負担が増加する')
p = doc.add_paragraph('　社会保険料は労使折半のため、新たに加入対象となる従業員が増えれば、その分の事業主負担も生じます。')
for run in p.runs:
    set_font(run)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = Pt(18)

add_bullet(doc, '従業員の手取り額が変わる', bold_part='従業員の手取り額が変わる')
p = doc.add_paragraph('　従業員側も保険料負担が生じるため、手取り額が減少します。特にパート・アルバイトスタッフへの丁寧な説明が求められます。')
for run in p.runs:
    set_font(run)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = Pt(18)

add_bullet(doc, '就業調整（いわゆる「壁」問題）が変化する', bold_part='就業調整（いわゆる「壁」問題）が変化する')
p = doc.add_paragraph('　賃金要件がなくなることで、従来の「8.8万円未満に抑える」という就業調整の意味が薄れます。労働時間ベースでの調整行動が増える可能性もあります。')
for run in p.runs:
    set_font(run)
p.paragraph_format.space_after = Pt(6)
p.paragraph_format.line_spacing = Pt(18)

add_body(doc,
    '対応にあたっては、社会保険労務士などの専門家に相談しながら、自社の状況に合った準備を進めることをおすすめします。',
    bold=False
)

# 保存
output_path = '/home/user/umb/社会保険適用拡大_解説.docx'
doc.save(output_path)
print(f'saved: {output_path}')
