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

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    run = p.add_run(text)
    if level == 1:
        set_font(run, bold=True, size=14, color=(0x1F, 0x49, 0x7D))
    elif level == 2:
        set_font(run, bold=True, size=11.5, color=(0x2E, 0x74, 0xB5))
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

def add_mixed(doc, parts):
    """parts: list of (text, bold)"""
    p = doc.add_paragraph()
    for text, bold in parts:
        run = p.add_run(text)
        set_font(run, bold=bold)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_bullet(doc, parts):
    """parts: list of (text, bold)"""
    p = doc.add_paragraph(style='List Bullet')
    for text, bold in parts:
        run = p.add_run(text)
        set_font(run, bold=bold)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_indent_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run(text)
    set_font(run)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = Pt(18)
    return p

def add_box(doc, text):
    """薄い背景色の注釈ボックス風段落"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_font(run)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = Pt(18)
    # 背景色
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'DEEAF1')
    pPr.append(shd)
    return p

# ===== 本文 =====

# タイトル
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = p_title.add_run('2026年度 厚生労働省の重点施策\n賃上げ・非正規労働者処遇改善支援を活用しましょう')
set_font(run_title, bold=True, size=15, color=(0x1F, 0x49, 0x7D))
p_title.paragraph_format.space_after = Pt(10)

# リード文
add_body(doc,
    '厚生労働省は、2026年度の労働行政における重点課題として「賃上げに向けた支援」と「非正規労働者の処遇改善」を掲げています。'
    'これらの施策には、事業主が積極的に活用できる助成金や支援策が含まれています。'
    '自社の取り組みと照らし合わせながら、使える制度を見落とさないようにしましょう。'
)

# 見出し1
add_heading(doc, '◎ 2026年度 企業向け施策の主なポイント', level=1)

add_body(doc, '今年度の重点施策は、大きく以下の２本柱です。')

add_bullet(doc, [('賃上げに取り組む企業への支援', True), ('（賃上げ支援助成金パッケージ・価格転嫁の徹底）', False)])
add_bullet(doc, [('非正規労働者の処遇改善に取り組む事業主への支援', True), ('（キャリアアップ助成金の活用促進）', False)])

# 見出し2
add_heading(doc, '◎ 賃上げ支援：助成金パッケージと価格転嫁の徹底', level=1)

add_body(doc,
    '賃上げに取り組む企業を後押しするため、厚生労働省は「賃上げ支援助成金パッケージ」の周知・活用促進を図っています。'
    'また、コストアップ分を適切に価格へ転嫁できるよう、取引適正化の徹底も合わせて推進されます。'
)

add_body(doc, '【主な助成金の例】')
add_bullet(doc, [('業務改善助成金', True), ('：生産性向上につながる設備投資等を行い、賃金を引き上げた中小企業・小規模事業者が対象', False)])
add_bullet(doc, [('人材確保等支援助成金', True), ('：雇用管理制度の整備・改善を通じて労働環境を向上させた企業が対象', False)])

add_box(doc,
    '【実務上のポイント】\n'
    '賃上げを実施する際は、助成金の申請要件（賃金引上げ率、対象労働者の範囲など）を事前に確認することが重要です。'
    '支給申請の期限を過ぎると受給できないため、実施前に最寄りの都道府県労働局やハローワークへ相談しておくとよいでしょう。'
)

# 見出し3
add_heading(doc, '◎ 非正規労働者の処遇改善：キャリアアップ助成金の活用', level=1)

add_body(doc,
    'パート・アルバイト・派遣社員など非正規労働者の待遇改善に取り組む事業主を支援する「キャリアアップ助成金」は、'
    '複数のコースから構成されており、自社の状況に応じた活用が可能です。'
)

add_body(doc, '【主なコース】')

add_bullet(doc, [('正社員化コース', True), ('：有期雇用労働者や短時間労働者を正社員・無期雇用労働者等に転換した場合に支給', False)])
add_indent_body(doc, '→ 1人あたり最大80万円（生産性要件を満たす場合）')

add_bullet(doc, [('賃金規定等改定コース', True), ('：非正規労働者の基本給などを3％以上増額改定した場合に支給', False)])
add_indent_body(doc, '→ 対象労働者の人数・増額率に応じて支給額が決定')

add_bullet(doc, [('賃金規定等共通化コース', True), ('：正規・非正規で別々だった賃金規定を統一・共通化した場合に支給', False)])

add_box(doc,
    '【実務上のポイント】\n'
    'キャリアアップ助成金は、就業規則や雇用契約書の整備が申請の前提となります。'
    '転換・処遇改善の「実施前」に計画を届け出る必要があるため、後から申請しようとしても受給できないケースがあります。'
    '取り組みを検討している場合は、早めに社会保険労務士や最寄りの労働局に確認することをおすすめします。'
)

# まとめ
add_heading(doc, '◎ まとめ：今すぐ確認しておきたいこと', level=1)

add_bullet(doc, [('賃上げを予定している', True), ('場合は、助成金の申請要件を実施前に確認する', False)])
add_bullet(doc, [('非正規労働者の正社員化・処遇改善を検討している', True), ('場合は、キャリアアップ助成金の対象コースを確認する', False)])
add_bullet(doc, [('就業規則・賃金規定が整備されているか', True), ('見直しておく', False)])
add_bullet(doc, [('わからない場合は', True), ('、都道府県労働局・ハローワーク・社会保険労務士へ早めに相談する', False)])

add_body(doc,
    '賃上げや処遇改善は、人材確保・定着にも直結します。助成金を賢く活用しながら、職場環境の改善を進めていきましょう。'
)

# 保存
output_path = '/home/user/umb/厚生労働省2026年度重点施策_解説.docx'
doc.save(output_path)
print(f'saved: {output_path}')
